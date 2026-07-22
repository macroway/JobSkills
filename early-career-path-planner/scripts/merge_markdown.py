#!/usr/bin/env python3
"""Merge the required career-plan sections into a reviewable Markdown draft."""

from __future__ import annotations

import argparse
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
PROTECTED_OUTPUTS = {"职业规划完整.md", "职业规划-精华版.md", "完整职业规划-用户版.md"}


def read_section(path: Path) -> str:
    try:
        text = path.read_text(encoding="utf-8").strip()
    except UnicodeDecodeError as exc:
        raise ValueError(f"文件不是 UTF-8：{path.name}") from exc
    if not text:
        raise ValueError(f"文件为空：{path.name}")
    return text


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path, help="包含七份分项文档的目录")
    parser.add_argument("--title", required=True, help="完整手册的一级标题")
    parser.add_argument("--output", default="08-职业规划完整初稿.md", help="输出文件名")
    args = parser.parse_args()

    folder = args.output_dir.expanduser().resolve()
    if not folder.is_dir():
        print(f"ERROR: 输出目录不存在：{folder}", file=sys.stderr)
        return 1

    missing = [name for name in SECTION_FILES if not (folder / name).is_file()]
    if missing:
        print(f"ERROR: 缺少分项文档：{', '.join(missing)}", file=sys.stderr)
        return 1

    title = args.title.strip()
    if not title:
        print("ERROR: 完整手册标题不能为空", file=sys.stderr)
        return 1

    destination = (folder / args.output).resolve()
    if destination.parent != folder:
        print("ERROR: 输出文件必须位于输出目录内", file=sys.stderr)
        return 1
    if destination.name in (*SECTION_FILES, *PROTECTED_OUTPUTS):
        print("ERROR: 输出文件名不能覆盖分项文档、完整成稿或精华版", file=sys.stderr)
        return 1

    try:
        sections = [read_section(folder / name) for name in SECTION_FILES]
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    content = f"# {title}\n\n" + "\n\n---\n\n".join(sections) + "\n"
    destination.write_text(content, encoding="utf-8", newline="\n")
    print(f"Created: {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
