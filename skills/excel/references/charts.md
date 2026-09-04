# 图表

openpyxl 支持柱状图、折线图、饼图、散点图等，写入后 Excel 可直接展示。

## 柱状图

```python
import openpyxl
from openpyxl.chart import BarChart, Reference

wb = openpyxl.Workbook()
ws = wb.active
for row in [("月份", "销量"), ("1月", 120), ("2月", 180), ("3月", 150)]:
    ws.append(row)

chart = BarChart()
chart.title = "月度销量"
data = Reference(ws, min_col=2, min_row=1, max_row=4)   # 数值列（含表头）
cats = Reference(ws, min_col=1, min_row=2, max_row=4)   # 类别列
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "E2")   # 图表锚点
wb.save("bar.xlsx")
```

## 折线图

```python
from openpyxl.chart import LineChart
chart = LineChart()
chart.title = "趋势"
...
```

## 饼图

```python
from openpyxl.chart import PieChart
chart = PieChart()
data = Reference(ws, min_col=2, min_row=1, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=4))
ws.add_chart(chart, "E2")
```

## 常见类型

| 类型 | 类名 |
|------|------|
| 柱状 | `BarChart` |
| 横向柱状 | `BarChart(type="bar")` |
| 折线 | `LineChart` |
| 饼图 | `PieChart` |
| 散点 | `ScatterChart` |
| 面积 | `AreaChart` |

## 注意

- `Reference` 的 `min_col`/`max_col` 指定数据列，`titles_from_data=True` 表示首行为标题。
- 图表锚点不能与数据重叠，一般放数据右侧空白列。
