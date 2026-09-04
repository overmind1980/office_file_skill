# 图片、表格与形状

## 插入图片

```python
from pptx import Presentation
from pptx.util import Inches
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.shapes.add_picture("chart.png", Inches(1), Inches(1), width=Inches(6))
prs.save("pic.pptx")
```

## 插入表格

```python
rows, cols = 3, 3
tbl = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), Inches(6), Inches(3)).table
tbl.cell(0, 0).text = "表头"
```

## 插入形状

```python
from pptx.enum.shapes import MSO_SHAPE
slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
```

## 注意

- 图片路径用绝对路径更稳。
- 大文件读取可用 `ppt.py read --md` 先看标题结构。
