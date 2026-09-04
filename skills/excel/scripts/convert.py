#!/usr/bin/env python3
"""convert.py — 表格格式互转 CLI（可移植）。

支持 .xls / .xlsx / .xlsm / .csv 之间的相互转换。

依赖策略（按需检查，缺失时给出安装提示）：
- 读 .xls 用原生 xlrd（含日期转换）；写 .xls 用原生 xlwt（pandas 2.x 已移除 xlwt writer）
- 读/写 .xlsx/.xlsm 用 openpyxl
- pandas 作为 .xlsx/.csv 读写主引擎
- 兜底方案：LibreOffice 命令行（见 references/xls.md），本脚本不依赖

用法：
  python3 convert.py in.xls --to xlsx [--out out.xlsx] [--sheet 名称或索引] [--encoding gbk]
  python3 convert.py in.xlsx --to xls
  python3 convert.py in.xlsm --to csv --sheet Sheet1
"""
import argparse
import json
import math
import os
import sys
from datetime import date, datetime


def need(mod, hint):
    try:
        __import__(mod)
    except ImportError:
        sys.exit(f"缺少依赖 {mod}：请执行 {hint}")


def detect_format(path):
    return os.path.splitext(path)[1].lower()


# --------------------------------------------------------------------------
# 读取 .xls（原生 xlrd，正确处理日期单元格）
# --------------------------------------------------------------------------
def _clean_header(header):
    seen = {}
    out = []
    for i, h in enumerate(header):
        if h is None or str(h).strip() == "":
            h = f"col{i + 1}"
        h = str(h)
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        out.append(h)
    return out


def _read_xls(path):
    import xlrd
    import pandas as pd
    wb = xlrd.open_workbook(path)
    out = {}
    for sh in wb.sheets():
        header, data = None, []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        row.append(xlrd.xldate_as_datetime(cell.value, wb.datemode))
                    except Exception:
                        row.append(cell.value)
                else:
                    row.append(cell.value)
            if r == 0:
                header = row
            else:
                data.append(row)
        if header is None:
            header = [f"col{i + 1}" for i in range(sh.ncols)]
        out[sh.name] = pd.DataFrame(data, columns=_clean_header(header))
    return out


# --------------------------------------------------------------------------
# 写入 .xls（原生 xlwt）
# --------------------------------------------------------------------------
def _xlwt_val(v):
    """把 pandas 单元格值转为 xlwt 可写类型。"""
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    try:
        pd = __import__("pandas")
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
    except Exception:
        pass
    if isinstance(v, datetime) or isinstance(v, date):
        return v
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float, str)):
        return v
    if hasattr(v, "item"):  # numpy 标量 -> python 原生
        return _xlwt_val(v.item())
    return str(v)


def _write_xls(path, sheets):
    import xlwt
    wb = xlwt.Workbook(encoding="utf-8")
    date_style = xlwt.XFStyle()
    date_style.num_format_str = "YYYY-MM-DD"
    for name, df in sheets.items():
        ws = wb.add_sheet(name[:31] or "Sheet")
        for c, col in enumerate(df.columns):
            ws.write(0, c, _xlwt_val(str(col)))
        for r in range(len(df)):
            for c, col in enumerate(df.columns):
                v = _xlwt_val(df.iloc[r, c])
                if isinstance(v, (datetime, date)):
                    ws.write(r + 1, c, v, date_style)
                else:
                    ws.write(r + 1, c, v)
    wb.save(path)


def main():
    p = argparse.ArgumentParser(prog="convert.py", description="表格格式互转 xls/xlsx/xlsm/csv")
    p.add_argument("input", help="源文件")
    p.add_argument("--to", required=True, choices=["xlsx", "xls", "csv"], help="目标格式")
    p.add_argument("--out", help="输出路径（默认同目录同名改扩展名）")
    p.add_argument("--sheet", help="仅转换指定 sheet（名称或从0开始的索引）")
    p.add_argument("--encoding", default="utf-8", help="csv 读写编码，默认 utf-8（老 xls 常用 gbk）")
    args = p.parse_args()

    if not os.path.exists(args.input):
        sys.exit(f"文件不存在: {args.input}")
    ext = detect_format(args.input)

    need("pandas", "pip3 install pandas")
    import pandas as pd

    # ---------- 读取源文件 ----------
    if ext == ".xls":
        need("xlrd", "pip3 install xlrd")
        sheets = _read_xls(args.input)
    elif ext in (".xlsx", ".xlsm"):
        need("openpyxl", "pip3 install openpyxl")
        xls = pd.ExcelFile(args.input, engine="openpyxl")
        sheets = {name: xls.parse(name) for name in xls.sheet_names}
    elif ext == ".csv":
        sheets = {os.path.splitext(os.path.basename(args.input))[0]:
                  pd.read_csv(args.input, encoding=args.encoding)}
    else:
        sys.exit(f"不支持的源格式: {ext}，仅支持 .xls/.xlsx/.xlsm/.csv")

    # ---------- 选择 sheet ----------
    names = list(sheets.keys())
    if args.sheet is not None:
        if args.sheet in sheets:
            selected = {args.sheet: sheets[args.sheet]}
        elif args.sheet.lstrip("-").isdigit():
            i = int(args.sheet)
            if i < 0 or i >= len(names):
                sys.exit(f"sheet 索引越界: {i}（共 {len(names)} 个）")
            selected = {names[i]: sheets[names[i]]}
        else:
            sys.exit(f"找不到 sheet: {args.sheet}，现有: {names}")
    else:
        selected = sheets

    if args.to == "csv":
        first = list(selected.keys())[0]
        if len(selected) > 1:
            print(f"警告: csv 只能输出单个 sheet，已取第一个 '{first}'", file=sys.stderr)
        df = selected[first]
        out = args.out or f"{os.path.splitext(args.input)[0]}.csv"
        df.to_csv(out, index=False, encoding=args.encoding)
        print(json.dumps({"out": out, "format": "csv", "sheet": first,
                          "rows": int(len(df))}, ensure_ascii=False))
        return

    out = args.out or f"{os.path.splitext(args.input)[0]}.{args.to}"

    if args.to == "xls":
        need("xlwt", "pip3 install xlwt")
        _write_xls(out, selected)
    else:  # xlsx
        need("openpyxl", "pip3 install openpyxl")
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for name, df in selected.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)

    print(json.dumps({"out": out, "format": args.to,
                      "sheets": [n[:31] for n in selected.keys()]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
