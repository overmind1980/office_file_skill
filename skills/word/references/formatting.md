# 样式与格式

`word.py` 不覆盖样式细节时，直接在 Python 中用 `python-docx` 处理。

## 字体与字号

```python
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

doc = Document()
p = doc.add_paragraph("中文文本")
run = p.add_run("这是加粗红色大字")
run.bold = True
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
run.font.name = "微软雅黑"
# 中文字体需设置 eastAsia
run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
doc.save("styled.docx")
```

## 段落对齐与缩进

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH

p = doc.add_paragraph("居中标题")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER   # LEFT/RIGHT/JUSTIFY

p2 = doc.add_paragraph("首行缩进两字符")
p2.paragraph_format.first_line_indent = Pt(28)
```

## 列表

```python
# 项目符号列表
doc.add_paragraph("项目一", style="List Bullet")
# 编号列表
doc.add_paragraph("步骤一", style="List Number")
```

## 分页

```python
doc.add_page_break()
```

## 页眉页脚

```python
sec = doc.sections[0]
sec.header.paragraphs[0].text = "页眉内容"
sec.footer.paragraphs[0].text = "第 1 页"
```

## 常用速查

| 目的 | 代码 |
|------|------|
| 加粗 | `run.bold = True` |
| 斜体 | `run.italic = True` |
| 下划线 | `run.underline = True` |
| 字号 | `run.font.size = Pt(12)` |
| 颜色 | `run.font.color.rgb = RGBColor(0,0,0)` |
| 中文字体 | `run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")` |
| 行距 | `p.paragraph_format.line_spacing = 1.5` |
