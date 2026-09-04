# 样式与格式

配合 `scripts/xlsx.py` 无法覆盖的样式需求时，可在 Python 中直接操作 openpyxl 样式对象。

## 完整示例

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active

# 表头样式
header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
for c in ws[1]:
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center", vertical="center")

# 边框
thin = Side(style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# 数字格式
ws["B2"].number_format = '#,##0.00'   # 千分位两位小数
ws["C2"].number_format = 'yyyy-mm-dd' # 日期

# 列宽/行高
ws.column_dimensions["A"].width = 18
ws.row_dimensions[1].height = 24

# 冻结首行（滚动时表头固定）
ws.freeze_panes = "A2"

# 自动筛选
ws.auto_filter.ref = ws.dimensions

# 合并单元格
ws.merge_cells("A1:C1")

wb.save("styled.xlsx")
```

## 常用样式速查

| 目的 | 代码 |
|------|------|
| 加粗红字 | `Font(bold=True, color="FF0000")` |
| 背景色 | `PatternFill("solid", start_color="FFFF00")` |
| 居中对齐 | `Alignment(horizontal="center", vertical="center")` |
| 文本换行 | `Alignment(wrap_text=True)` |
| 百分比 | `number_format = '0.00%'` |
| 千分位 | `number_format = '#,##0'` |
| 货币 | `number_format = '¥#,##0.00'` |

## 注意

- 颜色使用 6 位十六进制（无 `#`）。
- 样式需在 `fit` 到具体 `cell`/`row`/`column` 上才生效；修改已保存文件需重新 `save`。
