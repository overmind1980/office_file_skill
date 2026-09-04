# 表格

## 创建表格（word.py）

```bash
echo '[{"type":"table","rows":[["姓名","金额"],["张三",100],["李四",200]]}]' | python3 scripts/word.py write t.docx
```

## Python 直接创建

```python
from docx import Document
doc = Document()
table = doc.add_table(rows=2, cols=3)
table.cell(0, 0).text = "A1"
table.cell(1, 2).text = "C2"
doc.save("t.docx")
```

## 表格样式

```python
table.style = "Light Grid Accent 1"   # 内置样式名
```

## 合并单元格

```python
a = table.cell(0, 0)
b = table.cell(0, 1)
merged = a.merge(b)
merged.text = "合并后的单元格"
```

## 逐行填充数据

```python
data = [["表头1", "表头2"], ["a", "b"], ["c", "d"]]
table = doc.add_table(rows=len(data), cols=len(data[0]))
for ri, row in enumerate(data):
    for ci, val in enumerate(row):
        table.cell(ri, ci).text = str(val)
```

## 注意事项

- 合并单元格后读取 `cell.text` 会在合并区域重复出现相同文本。
- 样式名需与 Word 内置样式一致；不存在的样式名会报错或无效。
