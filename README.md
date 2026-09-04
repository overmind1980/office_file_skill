# 可移植技能包（Office 三件套）

一套符合「可移植 Skill 通用规范」的办公自动化技能集合，可无缝接入任意 LLM / Agent。

包含 3 个技能：**excel**（表格）、**word**（文档）、**ppt**（演示文稿）。每个技能均为自包含目录（`SKILL.md` + `scripts/` + `references/`），无绝对路径依赖，迁移即用。

## 目录结构

```
skills-pack/
├── README.md           # 本说明
├── requirements.txt    # Python 依赖清单
├── install.sh          # 一键安装依赖并自检
├── skills_loader.py    # 技能加载器（可复用，零第三方依赖）
└── skills/
    ├── excel/          # Excel 技能
    │   ├── SKILL.md
    │   ├── scripts/{xlsx.py, convert.py}
    │   └── references/{formulas,formatting,charts,xls}.md
    ├── word/           # Word 技能
    │   ├── SKILL.md
    │   ├── scripts/word.py
    │   └── references/{formatting,tables,images}.md
    └── ppt/            # PPT 技能
        ├── SKILL.md
        ├── scripts/ppt.py
        └── references/{layout,charts,images}.md
```

## 快速开始

```bash
# 1) 安装依赖并自检
./install.sh

# 2) 预览技能清单（验证可被发现）
python3 skills_loader.py ./skills
```

## 注册到智能体系统提示词

技能采用「渐进式披露」：系统提示词只注入**技能的 name + 用途**，模型按需用 `run_shell` 读取 `SKILL.md` 全文后调用脚本。集成只需两行：

```python
from skills_loader import build_skill_prompt

system_prompt += "\n\n" + build_skill_prompt(["./skills"])
```

`build_skill_prompt()` 会扫描指定目录，自动生成如下提示词片段：

```
可用技能（Skill）清单，任务匹配某项描述时优先使用对应技能：
- excel（目录 /abs/path/skills/excel）：读写与处理 Excel 工作簿……
- word（目录 /abs/path/skills/word）：创建、读取、编辑 Word 文档……
- ppt（目录 /abs/path/skills/ppt）：创建、读取、编辑 PowerPoint……

技能使用方法：先用 run_shell 读取该技能的 SKILL.md（cat "<技能目录>/SKILL.md"）……
```

> 若你的 Agent 已有技能发现逻辑，也可直接用 `discover_skills()` 拿到结构化数据自行拼接。

## 技能使用示例

### excel

```bash
python3 skills/excel/scripts/xlsx.py info 数据.xlsx
python3 skills/excel/scripts/xlsx.py read 数据.xlsx "Sheet1" --md
echo '[["姓名","金额"],["张三",100]]' | python3 skills/excel/scripts/xlsx.py write out.xlsx "报表"
python3 skills/excel/scripts/convert.py 旧表.xls --to xlsx --out 新表.xlsx
```

### word

```bash
python3 skills/word/scripts/word.py new 报告.docx --title "月度报告"
echo '[{"type":"heading","level":1,"text":"标题"},{"type":"table","rows":[["A","B"],["1","2"]]}]' \
  | python3 skills/word/scripts/word.py write 报告.docx
python3 skills/word/scripts/word.py read 报告.docx --md
```

### ppt

```bash
python3 skills/ppt/scripts/ppt.py new 汇报.pptx --title "季度汇报"
echo '[{"title":"封面","bullets":["作者"]},{"title":"业绩","bullets":["+20%"]}]' \
  | python3 skills/ppt/scripts/ppt.py build 汇报.pptx
python3 skills/ppt/scripts/ppt.py read 汇报.pptx --md
```

## 依赖

| 技能 | Python 库 | 用途 |
|------|-----------|------|
| excel | openpyxl / xlrd / xlwt / pandas | 读写 .xlsx/.xls、格式互转 |
| word | python-docx | 读写 .docx |
| ppt | python-pptx | 读写 .pptx |
| 通用 | （无） | skills_loader.py 仅用标准库 |

旧版 `.doc` / `.ppt` 二进制格式不受库直接支持，可先用 LibreOffice 转换：

```bash
libreoffice --headless --convert-to docx 文件.doc
libreoffice --headless --convert-to pptx 文件.ppt
```

## 环境要求

- Python 3.8+
- Linux / macOS / Windows 均可（脚本为纯 Python）
- 可选 LibreOffice（仅用于旧格式兜底转换）

## 如何新增自定义技能

1. 在 `skills/` 下新建目录，例如 `skills/myskill/`
2. 写一个带 frontmatter 的 `SKILL.md`：

```markdown
---
name: myskill
description: 这个技能做什么，何时使用。
---
# 技能正文（简洁，细节放 references/）
```

3. 可选添加 `scripts/`（可执行脚本）与 `references/`（细节文档）
4. 重新运行 `python3 skills_loader.py ./skills` 即自动识别，无需改任何代码

## 规范要点（可移植性）

- `SKILL.md` 必须有 `name`（小写字母/数字/连字符）与 `description`（说明何时使用）
- 技能自包含，不依赖文件绝对路径，脚本用相对路径调用
- 脚本默认输出 JSON，便于 Agent 程序化解析
- 主入口脚本名避免与第三方包同名（如用 `word.py` 而非 `docx.py`）

## 许可

MIT
