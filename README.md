# 绝区零队伍伤害计算器

通过 xlwings 直连本地 Excel，让 Excel 自己算公式，前端可视化展示结果。

## 原理

```
网页改配置 ──POST──→ Python 后端 ──xlwings──→ 本地 Excel
                                                      ↓
                                                Excel 自己算公式
                                                      ↓
网页渲染图表 ←──JSON── Python 后端 ←──xlwings── Excel 计算结果
```

**核心：Python 不算任何公式，只负责读写和传数据。Excel 是唯一的计算引擎。**

## 安装

```bash
pip install -r requirements.txt
```

需要：
- Python 3.8+
- Microsoft Excel（Windows）

## 使用

```bash
python local-server.py
```

浏览器打开 http://localhost:8081

## 文件说明

| 文件 | 说明 |
|------|------|
| `local-server.py` | Python 后端，Flask + xlwings |
| `index-local.html` | 前端页面，Chart.js 可视化 |
| `蕾米埃尔.xlsx` | Excel 计算表（核心逻辑） |

## 功能

- 7 个角色可选（维琳娜/简/蕾米埃尔/柚叶/爱丽丝/柏妮思/南宫羽）
- 命座、武器、精炼可调
- Boss 血量/防御/失衡/弱点/异常条系数可配
- 全队拐力 buff 表可编辑
- 伤害占比饼图、DPS 柱状图、时间分配、失衡积蓄
- 各角色单人伤害构成明细
- 分段伤害折线图（时间 vs 血量占比）
- 危局分数自动计算
