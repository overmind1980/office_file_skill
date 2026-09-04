# 图表

python-pptx 支持柱状图、折线图、饼图等原生图表。

## 柱状图

```python
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
chart_data = CategoryChartData()
chart_data.categories = ["一季度", "二季度", "三季度"]
chart_data.add_series("销量", (120, 180, 150))

x, y, cx, cy = Inches(1), Inches(1), Inches(6), Inches(4.5)
gframe = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)
prs.save("chart.pptx")
```

## 常见图表类型

| 类型 | XL_CHART_TYPE |
|------|---------------|
| 柱状 | COLUMN_CLUSTERED |
| 折线 | LINE |
| 饼图 | PIE |
| 条形 | BAR_CLUSTERED |

## 图表样式/标题

```python
chart = gframe.chart
chart.has_title = True
chart.chart_title.text_frame.text = "季度销量"
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
```

## 注意

- 图表数据用 `CategoryChartData` 组织。
- 图表文字不会被 `ppt.py read` 提取，需用本页 API 读取。
