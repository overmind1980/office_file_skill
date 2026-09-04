#!/usr/bin/env bash
# 可移植技能包安装脚本
set -e

echo "==> 安装 Python 依赖"
pip3 install -r "$(dirname "$0")/requirements.txt"

echo "==> 校验依赖"
python3 - <<'PY'
import openpyxl, xlrd, xlwt, pandas, docx, pptx
print("核心依赖 OK：openpyxl/xlrd/xlwt/pandas/python-docx/python-pptx")
PY

echo
echo "==> 校验技能包"
python3 "$(dirname "$0")/skills_loader.py" "$(dirname "$0")/skills"

echo
echo "安装完成。"
echo "使用方式：把本目录的 skills/ 与 skills_loader.py 复制（或软链）到你的 Agent 工程，"
echo "再在 Agent 代码中注入："
echo "    from skills_loader import build_skill_prompt"
echo "    system_prompt += \"\\n\\n\" + build_skill_prompt([\"./skills\"])"
