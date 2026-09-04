---
name: ppt
description: 创建、读取、编辑 PowerPoint 演示文稿（.pptx）。适用于生成幻灯片/汇报课件、读取每页标题与要点、搜索内容、设置版式、插入图表与图片。当用户提到 PPT、pptx、幻灯片、演示文稿、汇报/路演时使用本技能。
license: MIT
---

# PPT 技能

本技能提供 PowerPoint 演示文稿（.pptx）的创建、读取与内容提取能力，底层基于 Python `python-pptx`（可移植）。

## 统一入口

脚本 `scripts/ppt.py`，输出默认 JSON，便于程序化解析。

```bash
# 新建演示文稿（可选标题页）
python3 scripts/ppt.py new 汇报.pptx --title "季度汇报"

# 查看元信息（页数、每页标题）
python3 scripts/ppt.py info 汇报.pptx

# 读取每页标题与要点（JSON；--md 输出 Markdown）
python3 scripts/ppt.py read 汇报.pptx
python3 scripts/ppt.py read 汇报.pptx --md

# 搜索关键词
python3 scripts/ppt.py search 汇报.pptx "关键词"

# 按 JSON 结构生成 PPT（stdin 或 --data）
echo '[{"title":"封面","bullets":["作者：张三"]},{"title":"业绩","bullets":["收入 +20%","成本 -5%"]}]' \
  | python3 scripts/ppt.py build 汇报.pptx
```

## JSON 结构（build 命令）

```json
[
  {"title": "页面标题", "bullets": ["要点1", "要点2"], "layout": 1}
]
```

- `title`：页面标题；`bullets`：正文要点（列表）。
- `layout`：可选的版式索引，缺省第一页用标题版式（0），后续用“标题+内容”版式（1）。
- 常用版式：0=标题页，1=标题和内容，5=仅标题，6=空白。

## 依赖

- Python 3.8+
- `python-pptx`：`pip install python-pptx`（当前环境已就绪）
- 旧版 `.ppt` 不支持，先转格式：`libreoffice --headless --convert-to pptx 文件.ppt`

## 进阶参考

- 版式与布局（索引/标题/要点/文本框）：`references/layout.md`
- 图表（柱状/折线/饼图）：`references/charts.md`
- 图片/表格/形状：`references/images.md`

## 注意事项

- `python-pptx` 只处理 `.pptx`，`.ppt` 需先转换。
- `read` 只提取文本框文字，图片/图表中的文字不可提取（图表见 `references/charts.md`）。
- 生成后建议用 `read --md` 复核内容。
