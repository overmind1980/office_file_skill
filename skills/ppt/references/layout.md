# 版式与布局

## 常用版式索引（默认模板）

| 索引 | 版式 | 占位符 |
|------|------|--------|
| 0 | 标题幻灯片 | title + subtitle |
| 1 | 标题和内容 | title + body |
| 5 | 仅标题 | title |
| 6 | 空白 | 无 |

## 选择版式新建页

```python
from pptx import Presentation
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])   # 标题和内容
slide.shapes.title.text = "标题"
body = slide.placeholders[1]
tf = body.text_frame
tf.text = "第一点"
tf.add_paragraph().text = "第二点"
prs.save("out.pptx")
```

## 手动添加文本框

```python
from pptx.util import Inches, Pt
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
tb.text_frame.text = "自定义文本框"
```

## 设置项目符号级别

```python
tf = body.text_frame
p = tf.paragraphs[0]
p.level = 0   # 一级
p2 = tf.add_paragraph(); p2.text = "子项"; p2.level = 1
```

## 页面尺寸

```python
from pptx.util import Inches
prs.slide_width = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)
```

## 注意

- `slide.placeholders[1]` 不一定存在（取决于版式），访问前判断或 try/except。
- 标题用 `slide.shapes.title`（仅当版式含 title 占位符时非 None）。
