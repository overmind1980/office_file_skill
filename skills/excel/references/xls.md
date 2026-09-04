# .xls（Excel 97-2003）处理与转换

`.xls` 是旧版二进制格式（BIFF），与 `.xlsx`（XML/ZIP）不同，需用专用库处理。

## 能力对照

| 操作 | .xls（旧格式） | .xlsx（新格式） |
|------|---------------|----------------|
| 读取库 | `xlrd` | `openpyxl` |
| 写入库 | `xlwt`（原生） | `openpyxl` |
| pandas 读引擎 | `engine="xlrd"` | `engine="openpyxl"` |
| pandas 写引擎 | ❌ 2.x 已移除，用原生 `xlwt` | `engine="openpyxl"` |
| 行/列上限 | 65536 行 × 256 列 | 1048576 行 × 16384 列 |
| sheet 名长度 | ≤ 31 字符 | ≤ 31 字符 |

> ⚠️ pandas 2.x 已移除 xlwt 写入器，`df.to_excel(..., engine="xlwt")` 会报错；写 `.xls` 一律用原生 `xlwt`（见下）。

## 统一转换入口（推荐）

```bash
# .xls -> .xlsx（多 sheet 全保留）
python3 scripts/convert.py 旧表.xls --to xlsx --out 新表.xlsx

# .xlsx -> .xls（注意行/列数量别超 65536/256）
python3 scripts/convert.py 新表.xlsx --to xls

# 任意格式 -> csv
python3 scripts/convert.py 旧表.xls --to csv
python3 scripts/convert.py 新表.xlsx --to csv --sheet Sheet1

# csv -> xlsx / xls
python3 scripts/convert.py data.csv --to xlsx
python3 scripts/convert.py data.csv --to xls --encoding gbk
```

`--sheet` 支持 sheet 名称或从 0 开始的索引；`--encoding` 指定 CSV 编码（老文件常见 `gbk`）。

## 直接读写 .xls（Python）

### 读取（xlrd / pandas）

```python
import pandas as pd
# pandas 读 .xls 必须用 xlrd engine（xlrd 2.x 仅支持 .xls，不能读 .xlsx）
df = pd.read_excel("旧表.xls", sheet_name=0, engine="xlrd")

# 原生 xlrd
import xlrd
wb = xlrd.open_workbook("旧表.xls")
for sh in wb.sheets():
    print(sh.name, sh.nrows, sh.ncols)
    for r in range(sh.nrows):
        print(sh.row_values(r))
```

### 写入（原生 xlwt，pandas 2.x 无此 writer）

```python
import xlwt
wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("Sheet1")
for r, row in enumerate([["姓名", "金额"], ["张三", 100], ["李四", 200]]):
    for c, v in enumerate(row):
        ws.write(r, c, v)
wb.save("新表.xls")

# 从 DataFrame 手工写出（convert.py 内部即此逻辑）
def df_to_xls(path, name, df):
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(name[:31])
    for c, col in enumerate(df.columns):
        ws.write(0, c, str(col))
    for r in range(len(df)):
        for c, col in enumerate(df.columns):
            ws.write(r + 1, c, df.iloc[r, c])
    wb.save(path)
```

## LibreOffice 兜底方案（批量 / 服务器 / 无 pandas）

许多环境装有 LibreOffice，命令行转换最稳：

```bash
# 单个文件：.xls -> .xlsx
libreoffice --headless --convert-to xlsx 旧表.xls --outdir ./out

# 批量：目录内全部 .xls
libreoffice --headless --convert-to xlsx --outdir ./out *.xls

# 反向 .xlsx -> .xls
libreoffice --headless --convert-to xls 新表.xlsx

# csv -> xlsx
libreoffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" data.csv
```

未安装时可：`apt-get install -y libreoffice-calc --no-install-recommends`（体积较大，能不用就不用）。

## 常见坑

1. **编码**：老 .xls 里的中文常是 GBK；读 CSV 先试 `utf-8`，乱码再换 `--encoding gbk`。
2. **日期**：`.xls` 日期以序列号存储；只有被标记为日期单元格（ctype=3）才会由 `xlrd.xldate_as_datetime` 转成 `datetime`（`convert.py` 已处理）。已知坑：`xlwt` 写 `datetime` 时若不设日期格式样式，会存成无格式序列号，读回来是数字、Excel 里也显示数字：

   ```python
   style = xlwt.XFStyle()
   style.num_format_str = "YYYY-MM-DD"
   ws.write(r, c, dt, style)
   ```

3. **行/列上限**：`.xls` 最多 65536 行、256 列；`.xlsx -> .xls` 前先确认数据规模，超限会报错或截断。
4. **公式**：旧 .xls 中公式只能读出缓存值；xlrd 不解析公式，需要公式请先转 `.xlsx` 再交给 `scripts/xlsx.py` 处理。
5. **sheet 名**：长度 ≤31 字符，非法字符（`[]:*?/\\`）会被自动处理或报错。
6. **需要 LibreOffice 吗？** 有 pandas+xlrd+native xlwt 即可纯 Python 互转；LibreOffice 仅作兜底/批量场景。
