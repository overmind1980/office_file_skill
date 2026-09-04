---
name: word
description: 创建、读取、编辑 Word 文档（.docx）。适用于生成报告/文档、读取段落与表格内容、搜索文本、设置字体样式、插入表格与图片、分页。当用户提到 Word、docx、文档写作、报告生成、合同/简历等需求时使用本技能。
license: MIT
---

# Word 技能

本技能提供 Word 文档（.docx）的完整读写与生成能力，底层基于 Python `python-docx`（可移植）。

## 统一入口

脚本 `scripts/word.py`，输出默认 JSON，便于程序化解析。

```bash
# 新建文档
python3 scripts/word.py new 报告.docx --title "月度报告"

# 查看元信息（段落/表格/节数）
python3 scripts/word.py info 报告.docx

# 读取内容（JSON；--md 输出 Markdown）
python3 scripts/word.py read 报告.docx
python3 scripts/word.py read 报告.docx --md

# 搜索关键词
python3 scripts/word.py search 报告.docx "关键词"

# 按 JSON 结构生成/追加文档（stdin 或 --data）
echo '[{"type":"heading","level":1,"text":"标题"},{"type":"paragraph","text":"正文"},{"type":"table","rows":[["A","B"],["1","2"]]}]' \
  | python3 scripts/word.py write 报告.docx
```

## JSON 文档结构（write 命令）

```json
[
  {"type": "heading",    "level": 1,        "text": "一级标题"},
  {"type": "paragraph",  "text": "普通段落"},
  {"type": "table",      "rows": [["表头1","表头2"],["a","b"]]},
  {"type": "page_break"}
]
```

`write` 对已存在的文件是**追加**；不存在则新建。

## 依赖

- Python 3.8+
- `python-docx`：`pip install python-docx`（当前环境已就绪）
- 旧版 `.doc` 不支持，先转格式：`libreoffice --headless --convert-to docx 文件.doc`

## 进阶参考

- 样式与格式（字体/字号/颜色/对齐/列表/分页）：`references/formatting.md`
- 表格（创建/样式/合并单元格）：`references/tables.md`
- 图片（插入/尺寸/页眉页脚）：`references/images.md`

## 注意事项

- `python-docx` 只处理 `.docx`，`.doc` 需先转换。
- 表格合并单元格读取时可能出现文本重复（合并区域每个单元格返回相同文本），属正常现象。
- 生成后建议用 `read --md` 复核内容。
