"""
绝区零队伍伤害计算器 - 本地版 v2
xlwings 直连本地 Excel，Excel 自己算公式
pip install xlwings flask pywin32
python local-server.py
"""

import os, sys, json, time, threading
from pathlib import Path
from flask import Flask, jsonify, request, send_from_directory

try:
    import xlwings as xw
    import pythoncom
except ImportError:
    print("pip install xlwings flask pywin32")
    sys.exit(1)

pythoncom.CoInitialize()

app = Flask(__name__, static_folder='.')

# ============================================================
# Excel 控制器
# ============================================================

class ExcelController:
    def __init__(self, filepath):
        self.filepath = str(Path(filepath).resolve())
        self.app = None
        self.wb = None
        self.lock = threading.Lock()
        self._connect()

    def _connect(self):
        pythoncom.CoInitialize()
        print(f"打开 Excel: {self.filepath}")
        self.app = xw.App(visible=True, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = True
        self.wb = self.app.books.open(self.filepath)
        print(f"已打开: {self.wb.name}")

    def read(self, sheet, cell):
        with self.lock:
            try:
                pythoncom.CoInitialize()
                return self.wb.sheets[sheet].range(cell).value
            except Exception as e:
                print(f"读取失败 {sheet}!{cell}: {e}")
                return None

    def write(self, sheet, cell, value):
        with self.lock:
            try:
                pythoncom.CoInitialize()
                self.wb.sheets[sheet].range(cell).value = value
                return True
            except Exception as e:
                print(f"写入失败 {sheet}!{cell}: {e}")
                return False

    def calc(self):
        with self.lock:
            try:
                pythoncom.CoInitialize()
                self.app.calculate()
                time.sleep(0.2)
                return True
            except Exception as e:
                print(f"计算失败: {e}")
                return False

    def close(self):
        try:
            self.app.screen_updating = True
        except:
            pass

excel = None

def init_excel():
    global excel
    candidates = [
        Path(__file__).parent.parent / '蕾米埃尔.xlsx',
        Path(__file__).parent / '蕾米埃尔.xlsx',
        Path.cwd() / '蕾米埃尔.xlsx',
        Path.cwd().parent / '蕾米埃尔.xlsx',
    ]
    for p in candidates:
        if p.exists():
            excel = ExcelController(str(p))
            return
    print(f"找不到 蕾米埃尔.xlsx！搜索: {[str(p) for p in candidates]}")
    sys.exit(1)

# ============================================================
# 常量
# ============================================================

CHAR_ROWS = {
    '维':     {'row': 4, 'element': 'wind'},
    '简':     {'row': 5, 'element': 'phys'},
    '蕾':     {'row': 6, 'element': '光'},
    '柚叶':   {'row': 7, 'element': 'phys'},
    '爱丽丝': {'row': 8, 'element': 'phys'},
    '柏妮思': {'row': 9, 'element': 'fire'},
    '南宫羽': {'row': 10, 'element': 'ether'},
}

WEAPONS = {
    '维':     ['维专', '双生', '编译器'],
    '简':     ['简专', '双生', '编译器'],
    '蕾':     ['蕾专', '双生', '编译器'],
    '柚叶':   ['True', 'False'],
    '爱丽丝': ['爱专', '双生', '编译器'],
    '柏妮思': ['柏专', '双生', '编译器'],
    '南宫羽': ['专武', '双生', '编译器'],
}

# 元素名 → 中文（用于弱点显示）
ELEMENT_CN = {'wind': '风', 'phys': '物理', 'fire': '火', 'ether': '以太', '光': '光', 'elec': '电'}

# 全队拐力 buff 表头 (C1:O1)
BUFF_HEADERS = [
    {'col': 'C', 'key': 'atk',      'name': '攻击'},
    {'col': 'D', 'key': 'anomaly',  'name': '异常增伤'},
    {'col': 'E', 'key': 'mastery',  'name': '精通'},
    {'col': 'F', 'key': 'dmgbonus', 'name': '增伤'},
    {'col': 'G', 'key': 'accumres', 'name': '积蓄抗性'},
    {'col': 'H', 'key': 'crit',     'name': '强击暴击'},
    {'col': 'I', 'key': 'resist',   'name': '伤害抗性'},
    {'col': 'J', 'key': 'accumup',  'name': '积蓄提升'},
    {'col': 'K', 'key': 'defdown',  'name': '减防'},
    {'col': 'L', 'key': 'anomdmg',  'name': '强击增伤'},
    {'col': 'M', 'key': 'anomdef',  'name': '强击减防'},
    {'col': 'N', 'key': 'pen',      'name': '穿透'},
    {'col': 'O', 'key': 'disorder', 'name': '紊乱增伤'},
]

# 全队拐力中可编辑的行（角色buff + 危局buff + boss）
BUFF_ROWS = [
    {'row': 2, 'name': '维', 'label': '维琳娜'},
    {'row': 3, 'name': '柚', 'label': '柚叶'},
    {'row': 4, 'name': '简', 'label': '简'},
    {'row': 5, 'name': '蕾', 'label': '蕾米埃尔'},
    {'row': 6, 'name': '爱', 'label': '爱丽丝'},
    {'row': 7, 'name': '柏', 'label': '柏妮思'},
    {'row': 8, 'name': '南', 'label': '南宫羽'},
    {'row': 9, 'name': '危局buff', 'label': '危局Buff'},
    {'row': 10, 'name': 'boss', 'label': 'Boss'},
]

# boss 弱点元素（不含光）
WEAKNESS_ELEMENTS = ['风', '物理', '火', '以太']

@app.before_request
def _init_com():
    try:
        pythoncom.CoInitialize()
    except:
        pass

# ============================================================
# API
# ============================================================

@app.route('/')
def index():
    if Path('index-local.html').exists():
        return send_from_directory('.', 'index-local.html')
    return send_from_directory('.', 'index.html')

@app.route('/api/config', methods=['GET'])
def get_config():
    if not excel:
        return jsonify({'error': 'Excel 未连接'}), 500

    # 队伍
    team = []
    for name, info in CHAR_ROWS.items():
        r = info['row']
        team.append({
            'name': name,
            'element': info['element'],
            'selected': bool(excel.read('主操作表', f'D{r}')),
            'constellation': int(excel.read('主操作表', f'F{r}') or 0),
            'weapon': excel.read('主操作表', f'H{r}') or '',
            'weapons': WEAPONS.get(name, ['无']),
        })

    # Boss
    boss = {
        'hp': excel.read('主操作表', 'M23') or 205970837,
        'stun': excel.read('主操作表', 'K23') or 15486,
        'stunTime': excel.read('主操作表', 'K24') or 12,
        'stunVuln': excel.read('主操作表', 'T11') or 1.5,
        'effectiveTime': excel.read('主操作表', 'W8') or 180,
        'def': excel.read('主操作表', 'K28') or 60,
        'parryCount': int(excel.read('主操作表', 'T9') or 6),
        # 新增字段
        'shieldCount': int(excel.read('主操作表', 'T7') or 0),       # 秽盾数量
        'energyShield': int(excel.read('主操作表', 'T8') or 0),      # 能量盾数量
        'decibelOverflow': excel.read('主操作表', 'T10') or 1250,    # 单人喧响溢出
        'bossStunGift': excel.read('主操作表', 'W9') or 0,           # boss赠送失衡
        'anomalyCoeff': excel.read('主操作表', 'K25') or 1,          # 异常条系数
        'bossAnomalyCoeff': excel.read('主操作表', 'K26') or 1.1,    # 危局异常系数
        'initAccum': excel.read('主操作表', 'K27') or 3300,          # 初始积蓄
        'invincibleTime': excel.read('主操作表', 'W7') or 0,         # boss无敌时间
        # 弱点
        'weakness': {
            '风': excel.read('主操作表', 'F24') or '无',
            '物理': excel.read('主操作表', 'F25') or '无',
            '火': excel.read('主操作表', 'F26') or '无',
            '以太': excel.read('主操作表', 'F27') or '无',
        },
    }

    # 全队拐力 buff 表
    buffs = []
    for br in BUFF_ROWS:
        row_data = {'row': br['row'], 'name': br['name'], 'label': br['label']}
        for h in BUFF_HEADERS:
            val = excel.read('全队拐力', f"{h['col']}{br['row']}")
            row_data[h['key']] = val
        buffs.append(row_data)

    return jsonify({'team': team, 'boss': boss, 'buffs': buffs, 'buffHeaders': BUFF_HEADERS})

@app.route('/api/results', methods=['GET'])
def get_results():
    if not excel:
        return jsonify({'error': 'Excel 未连接'}), 500

    results = []
    for name, info in CHAR_ROWS.items():
        r = info['row']
        if not bool(excel.read('主操作表', f'D{r}')):
            continue
        time_val = excel.read('主操作表', f'I{r}') or 0
        damage = excel.read('主操作表', f'J{r}') or 0
        results.append({
            'name': name, 'element': info['element'],
            'constellation': int(excel.read('主操作表', f'F{r}') or 0),
            'weapon': excel.read('主操作表', f'H{r}') or '',
            'time': time_val, 'damage': damage,
            'damagePercent': excel.read('主操作表', f'K{r}') or 0,
            'stun': excel.read('主操作表', f'L{r}') or 0,
            'stunPercent': excel.read('主操作表', f'M{r}') or 0,
            'accumulation': excel.read('主操作表', f'N{r}') or 0,
            'accumPercent': excel.read('主操作表', f'O{r}') or 0,
            'decibel': excel.read('主操作表', f'P{r}') or 0,
            'energy': excel.read('主操作表', f'Q{r}') or 0,
            'dps': damage / time_val if time_val > 0 else 0,
        })

    total_damage = excel.read('主操作表', 'J13') or 0
    total_score = excel.read('主操作表', 'M90') or 0
    hp_percent = excel.read('主操作表', 'L13') or 0

    # 危局分数
    score_rows = []
    for sr in range(83, 90):
        hp_range = excel.read('主操作表', f'C{sr}') or ''
        bars = 0
        if '-' in str(hp_range):
            parts = str(hp_range).split('-')
            if len(parts) == 2:
                try: bars = int(parts[0]) - int(parts[1]) + 1
                except: pass
        score_rows.append({
            'range': hp_range, 'bars': bars,
            'perBar': excel.read('主操作表', f'D{sr}') or 0,
            'score': excel.read('主操作表', f'M{sr}') or 0,
        })

    # 各角色单人伤害明细
    char_detail = {}
    detail_defs = {
        '维': {'sheet':'资源表', 'items':[(69,'直伤'),(70,'风化'),(71,'异放'),(72,'乱流')], 'val_col':'C', 'pct_col':'D', 'total_row':74, 'total_col':'C'},
        '蕾': {'sheet':'资源表', 'items':[(76,'直伤'),(77,'流明')], 'val_col':'N', 'pct_col':'O', 'total_row':79, 'total_col':'N'},
        '柏妮思': {'sheet':'资源表', 'items':[(69,'燃烧'),(70,'紊乱'),(71,'异放'),(72,'直伤')], 'val_col':'AB', 'pct_col':None, 'total_row':73, 'total_col':'AB'},
        '爱丽丝': {'sheet':'资源表', 'items':[(53,'直伤'),(54,'强击'),(55,'紊乱'),(56,'dot')], 'val_col':'AK', 'pct_col':'AL', 'total_row':58, 'total_col':'AK'},
        '南宫羽': {'sheet':'资源表', 'items':[(81,'侵蚀'),(82,'紊乱'),(83,'直伤'),(84,'异放'),(85,'极性紊乱')], 'val_col':'BC', 'pct_col':'BE', 'total_row':87, 'total_col':'BD'},
        '简': {'sheet':'资源表', 'items':[(40,'直伤'),(41,'强击'),(42,'紊乱')], 'val_col':'BK', 'pct_col':'BL', 'total_row':44, 'total_col':'BK'},
    }
    for name, info in CHAR_ROWS.items():
        r = info['row']
        if not bool(excel.read('主操作表', f'D{r}')):
            continue
        dd = detail_defs.get(name)
        if not dd:
            continue
        total = excel.read(dd['sheet'], f"{dd['total_col']}{dd['total_row']}") or 0
        items = []
        for row, label in dd['items']:
            val = excel.read(dd['sheet'], f"{dd['val_col']}{row}") or 0
            pct = 0
            if dd['pct_col']:
                pct = excel.read(dd['sheet'], f"{dd['pct_col']}{row}") or 0
            elif total > 0:
                pct = val / total
            items.append({'name': label, 'value': val, 'percent': pct})
        char_detail[name] = {'items': items, 'total': total}

    # 分段伤害（时间 vs 血量占比）
    segment_data = []
    for r in range(7, 14):
        t = excel.read('主操作表', f'AA{r}')
        hp = excel.read('主操作表', f'AB{r}')
        if t is not None and hp is not None and hp is not False:
            segment_data.append({'time': t, 'hpPercent': hp})

    # 喧响 & 能量
    decibel_energy = []
    for name, info in CHAR_ROWS.items():
        r = info['row']
        if not bool(excel.read('主操作表', f'D{r}')):
            continue
        decibel_energy.append({
            'name': name,
            'decibel': excel.read('主操作表', f'P{r}') or 0,
            'energy': excel.read('主操作表', f'Q{r}') or 0,
        })

    return jsonify({
        'results': results,
        'totalDamage': total_damage,
        'totalScore': total_score,
        'hpPercent': hp_percent,
        'effectiveTime': excel.read('主操作表', 'W8') or 180,
        'charDetail': char_detail,
        'decibelEnergy': decibel_energy,
        'segmentData': segment_data,
    })

@app.route('/api/update', methods=['POST'])
def update_config():
    if not excel:
        return jsonify({'error': 'Excel 未连接'}), 500

    data = request.json
    if not data:
        return jsonify({'error': 'No data'}), 400

    try:
        # 队伍配置
        if 'team' in data:
            for cfg in data['team']:
                name = cfg.get('name', '')
                if name not in CHAR_ROWS:
                    continue
                r = CHAR_ROWS[name]['row']
                excel.write('主操作表', f'D{r}', cfg.get('selected', False))
                excel.write('主操作表', f'F{r}', cfg.get('constellation', 0))
                excel.write('主操作表', f'H{r}', cfg.get('weapon', ''))

        # Boss 配置
        if 'boss' in data:
            b = data['boss']
            def bw(key, cell, default=None):
                if key in b and b[key] is not None:
                    excel.write('主操作表', cell, b[key])
            bw('hp', 'M23')
            bw('stun', 'K23')
            bw('stunTime', 'K24')
            bw('stunVuln', 'T11')
            bw('def', 'K28')
            bw('parryCount', 'T9')
            bw('shieldCount', 'T7')
            bw('energyShield', 'T8')
            bw('decibelOverflow', 'T10')
            bw('bossStunGift', 'W9')
            bw('anomalyCoeff', 'K25')
            bw('bossAnomalyCoeff', 'K26')
            bw('initAccum', 'K27')
            bw('invincibleTime', 'W7')

            # 弱点
            if 'weakness' in b:
                weakness_map = {'风': 'F24', '物理': 'F25', '火': 'F26', '以太': 'F27'}
                resist_vals = {'无': ('0', '0'), '弱点': ('0.2', '0.2'), '抗性': ('-0.4', '-0.2')}
                for elem, cell in weakness_map.items():
                    val = b['weakness'].get(elem, '无')
                    excel.write('主操作表', cell, val)
                    fv, gv = resist_vals.get(val, ('0', '0'))
                    # G列=伤害抗性, H列=其他抗性
                    col_num = int(cell[1:])  # row number
                    excel.write('主操作表', f'G{col_num}', float(fv))
                    excel.write('主操作表', f'H{col_num}', float(gv))

        # 全队拐力 buff 表
        if 'buffs' in data:
            for row_data in data['buffs']:
                row = row_data.get('row')
                if not row:
                    continue
                for h in BUFF_HEADERS:
                    key = h['key']
                    if key in row_data and row_data[key] is not None:
                        excel.write('全队拐力', f"{h['col']}{row}", row_data[key])

        # Excel 自己算
        excel.calc()
        return jsonify({'status': 'ok'})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/debug', methods=['GET'])
def debug():
    sheet = request.args.get('sheet', '主操作表')
    cell = request.args.get('cell', 'A1')
    return jsonify({'sheet': sheet, 'cell': cell, 'value': excel.read(sheet, cell)})

# ============================================================
if __name__ == '__main__':
    print("=" * 50)
    print("绝区零队伍伤害计算器 - 本地版 v2")
    print("=" * 50)
    init_excel()
    print(f"\n浏览器打开: http://localhost:8081\n")
    try:
        app.run(host='127.0.0.1', port=8081, debug=False, use_reloader=False, threaded=False)
    finally:
        if excel: excel.close()
