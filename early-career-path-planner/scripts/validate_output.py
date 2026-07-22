#!/usr/bin/env python3
"""Validate the required files and weekday action plan in a career-plan folder."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


SECTION_FILES = (
    "01-职业画像.md",
    "02-岗位检索与职业地图.md",
    "03-能力差距矩阵.md",
    "04-成长路线.md",
    "05-每周行动清单.md",
    "06-求职策略与备选路径.md",
    "07-来源与假设.md",
)
FINAL_FILE = "职业规划完整.md"
USER_FILE = "职业规划-精华版.md"
USER_REQUIRED_HEADINGS = (
    "## 当前状态与已有基础",
    "### 当前状态",
    "### 简历事实",
    "### 用户表达的偏好",
    "## 先说结论",
    "## 你提出的问题，我的回答",
    "## 更适合你的岗位",
    "## 接下来怎么投",
    "## 接下来 12 周先做什么",
    "## 需要补的能力与作品",
    "## 现在就做的三件事",
)
USER_OPENING_HEADINGS = (
    "## 当前状态与已有基础",
    "### 当前状态",
    "### 简历事实",
    "### 用户表达的偏好",
    "## 先说结论",
)
PREFERENCE_HEADING = "### 用户表达的偏好"
REQUIRED_HEADINGS = (
    "## 职业画像",
    "## 岗位检索与职业地图",
    "## 能力差距矩阵",
    "## 成长路线",
    "## 每周行动清单",
    "## 求职策略与备选路径",
    "## 来源与假设",
)
WEEK_PATTERN = re.compile(r"^###\s*第\s*\d+\s*周.*?(?=^###\s*第\s*\d+\s*周|\Z)", re.MULTILINE | re.DOTALL)
USER_FORBIDDEN_PATTERNS = (
    ("来源编号", re.compile(r"\b(?:POST|LOCAL|RHY|RESUME|AUTH|ORG|JD)-[A-Za-z0-9]+-\d+\b")),
    ("逐周任务汇总", re.compile(r"(?m)^本周总工时\s*[:：]")),
    ("逐日任务表", re.compile(r"(?m)^\|\s*周[一二三四五]\s*\|")),
)


def report(errors: list[str], warnings: list[str]) -> int:
    for item in errors:
        print(f"ERROR: {item}")
    for item in warnings:
        print(f"WARNING: {item}")
    if errors:
        return 1
    print("OK: 输出目录结构、完整成稿、精华版和按日行动计划均已通过检查。")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()

    folder = args.output_dir.expanduser().resolve()
    errors: list[str] = []
    warnings: list[str] = []
    if not folder.is_dir():
        return report([f"输出目录不存在：{folder}"], warnings)

    if not re.fullmatch(r".+-.+-.+-\d{8}-\d{6}", folder.name):
        errors.append("文件夹名称必须为 <姓名-当前职级-主方向岗位名称-YYYYMMDD-HHmmss>")

    for name in (*SECTION_FILES, FINAL_FILE, USER_FILE):
        path = folder / name
        if not path.is_file():
            errors.append(f"缺少文件：{name}")
            continue
        try:
            if not path.read_text(encoding="utf-8").strip():
                errors.append(f"文件为空：{name}")
        except UnicodeDecodeError:
            errors.append(f"文件不是 UTF-8：{name}")

    weekly_path = folder / "05-每周行动清单.md"
    if weekly_path.is_file():
        weekly = weekly_path.read_text(encoding="utf-8")
        weeks = list(WEEK_PATTERN.finditer(weekly))
        if not weeks:
            errors.append("每周行动清单缺少“### 第 N 周”区块")
        if re.search(r"本周总工时\s*[:：]|\|\s*日期\s*\|\s*工时\s*\||^\|\s*周[一二三四五]\s*\|\s*\d+\s*小时\s*\|", weekly, re.MULTILINE):
            errors.append("每周行动清单不应包含工时列、小时数或周总工时")
        for index, match in enumerate(weeks, start=1):
            block = match.group(0)
            for weekday in ("周一", "周二", "周三", "周四", "周五"):
                row = re.compile(rf"^\|\s*{weekday}\s*\|", re.MULTILINE)
                if not row.search(block):
                    errors.append(f"第 {index} 周缺少“{weekday}”任务行")
            if re.search(r"^\|\s*周[六日]\s*\|", block, re.MULTILINE):
                errors.append(f"第 {index} 周包含周末任务")

    final_path = folder / FINAL_FILE
    if final_path.is_file():
        final = final_path.read_text(encoding="utf-8")
        previous_heading_index = -1
        for heading in REQUIRED_HEADINGS:
            heading_index = final.find(heading)
            if heading_index == -1:
                errors.append(f"完整成稿缺少章节：{heading}")
            elif heading_index <= previous_heading_index:
                errors.append(f"完整成稿章节顺序错误：{heading}")
            else:
                previous_heading_index = heading_index
        if PREFERENCE_HEADING not in final:
            errors.append("完整成稿的职业画像缺少“用户表达的偏好”小节")
        for phrase in ("希望这对您有帮助", "当然！", "一定！"):
            if phrase in final:
                warnings.append(f"完整成稿仍含聊天痕迹：{phrase}")

    user_path = folder / USER_FILE
    if user_path.is_file():
        user_text = user_path.read_text(encoding="utf-8")
        if not user_text.startswith("#"):
            errors.append("精华版必须以一级标题开始")
        for heading in USER_REQUIRED_HEADINGS:
            if heading not in user_text:
                errors.append(f"精华版缺少章节：{heading}")
        previous_heading_index = -1
        for heading in USER_OPENING_HEADINGS:
            heading_index = user_text.find(heading)
            if heading_index == -1:
                continue
            if heading_index <= previous_heading_index:
                errors.append(f"精华版开头结构顺序错误：{heading}")
            previous_heading_index = heading_index
        first_second_level = re.search(r"^##\s+[^#]", user_text, re.MULTILINE)
        if first_second_level and not user_text[first_second_level.start():].startswith("## 当前状态与已有基础"):
            errors.append("精华版的第一个一级部分必须是：当前状态与已有基础")
        for marker in ("http://", "https://", "skills/early-career-path-planner/"):
            if marker in user_text:
                errors.append(f"精华版不应包含报告实现细节：{marker}")
        for label, pattern in USER_FORBIDDEN_PATTERNS:
            if pattern.search(user_text):
                errors.append(f"精华版不应包含报告实现细节：{label}")
        for phrase in ("希望这对您有帮助", "当然！", "一定！"):
            if phrase in user_text:
                warnings.append(f"精华版仍含聊天痕迹：{phrase}")

    return report(errors, warnings)


if __name__ == "__main__":
    raise SystemExit(main())
