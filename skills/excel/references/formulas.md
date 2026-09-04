# 公式与函数

openpyxl 中写入公式即把单元格值设为以 `=` 开头的字符串。`read`（默认 `data_only=False`）会返回公式文本；要获取计算结果需要 "数据模式" 读取。

## 常用公式

| 类别 | 示例 |
|------|------|
| 求和 | `=SUM(A1:A10)` |
| 条件求和 | `=SUMIF(A:A,">100")` |
| 平均值 | `=AVERAGE(B2:B20)` |
| 计数 | `=COUNTA(A2:A20)` |
| 条件统计 | `=COUNTIF(C:C,"完成")` |
| 查找 | `=VLOOKUP(E2, A:B, 2, FALSE)` |
| 条件 | `=IF(A2>0,"正","负")` |
| 拼接 | `=A2 & "-" & B2`，或 `=TEXTJOIN(",",TRUE,A2:C2)` |
| 日期 | `=TODAY()`、`=YEAR(A2)`、`=DATEDIF(A2,B2,"d")` |

## Python 写入公式（脚本内）

```python
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws["A1"] = 10
ws["A2"] = 20
ws["A3"] = "=SUM(A1:A2)"   # 写入公式，Excel 打开时自动计算
wb.save("f.xlsx")
```

## 读取计算后的结果

```python
from openpyxl import load_workbook
# 只返回上次 Excel 保存时缓存的计算结果；公式未被计算过则为 None
wb = load_workbook("f.xlsx", data_only=True)
print(wb.active["A3"].value)
```

## 通过 xlsx.py 写入公式

公式只是普通字符串，直接放进 JSON 数组即可：

```bash
echo '[[10,20],["=SUM(A1:A2)"]]' | python3 scripts/xlsx.py write f.xlsx "Sheet1"
```

## 注意

- 用 Python 计算的公式需手动刷新，或先用 LibreOffice 重算并保存：
  `libreoffice --headless --convert-to xlsx --calc f.xlsx`（另存为再打开）。
- 区域引用不要含 `$` 时照写：`=SUM($A$1:$A$10)` 表示绝对引用。
