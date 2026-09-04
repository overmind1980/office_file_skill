---
name: excel
description: 读写与处理 Excel 工作簿（.xlsx/.xlsm/.xls）。适用于创建、读取、搜索、编辑表格数据，生成报表，处理多 sheet 工作簿，写入公式、格式与图表，以及 xls/xlsx/csv 之间的格式转换。当用户提到 Excel、xlsx、xls、表格/电子表格、报表导出与格式转换等需求时使用本技能。
license: MIT
---

# Excel 技能

本技能提供 Excel 工作簿（.xlsx / .xlsm / .xls）的完整读写、处理与格式转换能力，底层基于 Python `openpyxl` / `xlrd` / `xlwt` / `pandas`（可移植，仅需标准 Python 运行时）。

## 统一入口

本技能提供两个可执行脚本，输出均为 JSON/Markdown，便于程序化解析：

- `scripts/xlsx.py` — 新版 `.xlsx/.xlsm` 的读写、搜索、写入公式等
- `scripts/convert.py` — `.xls / .xlsx / .xlsm / .csv` 互转

### xlsx.py（.xlsx/.xlsm 处理）

```bash
# 查看工作簿元信息（sheet 列表及行列数）
python3 scripts/xlsx.py info 数据.xlsx

# 仅列出 sheet 名称
python3 scripts/xlsx.py sheets 数据.xlsx

# 读取 sheet（默认 JSON，整表）
python3 scripts/xlsx.py read 数据.xlsx "Sheet1"
python3 scripts/xlsx.py read 数据.xlsx "Sheet1" --range A1:C20
python3 scripts/xlsx.py read 数据.xlsx "Sheet1" --md      # Markdown 表格
python3 scripts/xlsx.py read 数据.xlsx "Sheet1" --csv     # CSV

# 搜索内容
python3 scripts/xlsx.py search 数据.xlsx "关键词"

# 写入数据（JSON 二维数组，stdin 或 --data）
echo '[["姓名","金额"],["张三",100],["李四",200]]' | python3 scripts/xlsx.py write out.xlsx "报表"

# 新建工作簿
python3 scripts/xlsx.py new out.xlsx
```

### convert.py（格式互转）

```bash
# .xls -> .xlsx（多 sheet 全保留）
python3 scripts/convert.py 旧表.xls --to xlsx --out 新表.xlsx

# .xlsx -> .xls（注意别超 65536 行 × 256 列）
python3 scripts/convert.py 新表.xlsx --to xls

# 任意格式 -> csv
python3 scripts/convert.py 旧表.xls --to csv
python3 scripts/convert.py 新表.xlsx --to csv --sheet Sheet1

# csv -> xlsx / xls
python3 scripts/convert.py data.csv --to xlsx
python3 scripts/convert.py data.csv --to xls --encoding gbk   # 老中文 CSV 常用 gbk
```

每个子命令都支持 `-h` 查看参数。

## 依赖

- Python 3.8+（当前环境已就绪）
- `openpyxl`：读写 `.xlsx/.xlsm` — `pip install openpyxl`
- `xlrd`：读取 `.xls`（只读） — `pip install xlrd`
- `xlwt`：写入 `.xls` — `pip install xlwt`
- `pandas`：`convert.py` 的主引擎（推荐） — `pip install pandas`
- 兜底：`libreoffice --headless --convert-to ...`（批量转换/服务器场景，见 `references/xls.md`）

遇到缺依赖先按提示 `pip3 install <包>` 再重试；不要用 `sudo`。

## 标准操作流程

1. **了解结构**：先 `info` 或 `sheets` 查看 sheet 名与行列规模（.xls 先用 `convert.py` 转成 .xlsx 更方便）。
2. **读取数据**：`read` 加 `--range` 限定区域，避免大表全量载入。
3. **搜索定位**：`search` 快速找到关键字所在 sheet 与单元格。
4. **写入/生成**：把数据整理成 JSON 二维数组传入 `write`；已存在文件默认**追加**到同名 sheet 末尾，加 `--overwrite` 才替换。
5. **格式转换**：`.xls` 相关一律走 `convert.py`，多 sheet 自动全保留；CSV 注意 `--encoding`。
6. **验证**：写完后用 `read --md` 或 `info` 复核结果。

## 进阶参考

- `.xls` 旧格式处理与转换（xlrd/xlwt/LibreOffice/编码坑）：`references/xls.md`
- 公式与函数：`references/formulas.md`
- 样式与格式（字体/边框/数字格式/列宽/冻结）：`references/formatting.md`
- 图表（柱状/折线/饼图）：`references/charts.md`

## 注意事项

- 写操作会覆盖文件的其他改动，先备份原文件再批量写入。
- 公式单元格 `read` 默认返回公式文本（`data_only=False`）；需计算值参考 `references/formulas.md`。
- `.xlsx -> .xls` 前确认数据规模：`.xls` 上限 65536 行 × 256 列、sheet 名 ≤31 字符。
- 老 `.xls`/CSV 中文乱码时，用 `--encoding gbk` 重试。
