# 图片、页眉页脚

## 插入图片

```python
from docx import Document
from docx.shared import Inches

doc = Document()
doc.add_paragraph("下面是图片：")
doc.add_picture("chart.png", width=Inches(4), height=Inches(3))
doc.save("img.docx")
```

支持 PNG/JPG/GIF/BMP 等常见格式。

## 尺寸单位

```python
from docx.shared import Inches, Cm, Pt, Emu
doc.add_picture("a.png", width=Cm(8))
```

## 设置文档方向与页边距

```python
from docx.shared import Cm
sec = doc.sections[0]
sec.orientation = 1          # 0=纵向 1=横向（结合下面互换宽高）
sec.page_width, sec.page_height = sec.page_height, sec.page_width
sec.top_margin = Cm(2.5)
```

## 页眉页脚

```python
sec = doc.sections[0]
sec.header.paragraphs[0].text = "机密文档"
sec.footer.paragraphs[0].text = "第 1 页"
```

## 注意

- 图片路径用绝对路径更稳。
- 文档较大时用 `word.py read --max` 限制读取块数。
