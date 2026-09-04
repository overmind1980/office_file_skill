#!/usr/bin/env python3
"""skills_loader.py — 可移植技能(Skill)加载器（零第三方依赖）。

让任意 LLM Agent 快速接入「可移植技能」目录：扫描 skills/ 下每个含
SKILL.md 的子目录，解析其 YAML frontmatter(name/description)，生成
可注入系统提示词的技能清单（渐进式披露：仅注入名称与用途，详情由
模型按需读取 SKILL.md 全文）。

用法示例（集成到你的 Agent）：
    from skills_loader import build_skill_prompt
    system_prompt += "\n\n" + build_skill_prompt(["./skills"])

也可命令行直接预览：
    python3 skills_loader.py ./skills
"""
import os
import re


def parse_skill_frontmatter(path):
    """解析 SKILL.md 的 YAML frontmatter，返回 dict 或 None。"""
    try:
        with open(path, encoding="utf-8") as f:
            text = f.read()
    except OSError:
        return None
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n?", text, re.DOTALL)
    if not m:
        return None
    fm = {}
    for line in m.group(1).splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            fm[k.strip()] = v.strip()
    return fm


def discover_skills(skills_dirs=None):
    """扫描技能根目录，返回 [{'name','dir','description'}]。"""
    if skills_dirs is None:
        skills_dirs = ["./skills"]
    found = []
    for root in skills_dirs:
        if not os.path.isdir(root):
            continue
        for entry in sorted(os.listdir(root)):
            d = os.path.join(root, entry)
            md = os.path.join(d, "SKILL.md")
            if os.path.isdir(d) and os.path.isfile(md):
                fm = parse_skill_frontmatter(md)
                if fm and fm.get("name"):
                    found.append({
                        "name": fm["name"],
                        "dir": os.path.abspath(d),
                        "description": fm.get("description", ""),
                    })
    return found


def build_skill_prompt(skills_dirs=None):
    """生成可注入系统提示词的技能清单文本。"""
    skills = discover_skills(skills_dirs)
    if not skills:
        return "可用技能：无。"
    lines = "\n".join(
        f"- {s['name']}（目录 {s['dir']}）：{s['description']}" for s in skills
    )
    return (
        "可用技能（Skill）清单，任务匹配某项描述时优先使用对应技能：\n"
        + lines + "\n\n"
        "技能使用方法：先用 run_shell 读取该技能的 SKILL.md（cat \"<技能目录>/SKILL.md\"）"
        "获取完整用法与命令；再按其指引调用该技能 scripts/ 下的脚本，必要时查阅 references/。"
        "调用脚本时请使用上表中技能目录的绝对路径。"
    )


if __name__ == "__main__":
    import sys
    dirs = sys.argv[1:] if len(sys.argv) > 1 else None
    print(build_skill_prompt(dirs))
