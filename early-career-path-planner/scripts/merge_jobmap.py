#!/usr/bin/env python3
"""Merge per-level jobmap files into one file per job title."""

from __future__ import annotations

import argparse
import re
from collections import defaultdict
from pathlib import Path

LEVEL_SUFFIXES = ("负责人", "校招", "初级", "中级", "高级", "专家")
LEVEL_ORDER = {"校招": 0, "初级": 1, "中级": 2, "高级": 3, "专家": 4, "负责人": 5}


def parse_filename(path: Path) -> tuple[str, str] | None:
    stem = path.stem
    for suffix in LEVEL_SUFFIXES:
        marker = f"-{suffix}"
        if stem.endswith(marker):
            return stem[: -len(marker)], suffix
    return None


def strip_title(content: str) -> str:
    lines = content.splitlines()
    if lines and lines[0].startswith("# ") and "岗位职责描述" in lines[0]:
        body = "\n".join(lines[1:])
        return body.lstrip("\n")
    return content


def merge_job(job_name: str, level_files: dict[str, Path]) -> str:
    title = f"# {job_name}｜岗位职责描述"
    sections: list[str] = [title, ""]
    for level in sorted(level_files, key=lambda item: LEVEL_ORDER.get(item, 99)):
        body = strip_title(level_files[level].read_text(encoding="utf-8")).strip()
        sections.extend([f"## {level}", "", body, ""])
    return "\n".join(sections).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "jobmap_dir",
        type=Path,
        nargs="?",
        default=Path(__file__).resolve().parents[1] / "references" / "jobmap",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    jobmap_dir = args.jobmap_dir.expanduser().resolve()
    if not jobmap_dir.is_dir():
        raise SystemExit(f"目录不存在：{jobmap_dir}")

    grouped: dict[str, dict[str, Path]] = defaultdict(dict)
    skipped: list[str] = []

    for path in sorted(jobmap_dir.glob("*.md")):
        parsed = parse_filename(path)
        if not parsed:
            skipped.append(path.name)
            continue
        job_name, level = parsed
        grouped[job_name][level] = path

    if skipped:
        print(f"跳过非职级文件 {len(skipped)} 个：{', '.join(skipped[:5])}" + (" ..." if len(skipped) > 5 else ""))

    created = 0
    removed = 0
    for job_name in sorted(grouped):
        out_path = jobmap_dir / f"{job_name}.md"
        merged = merge_job(job_name, grouped[job_name])
        if args.dry_run:
            print(f"[dry-run] {out_path.name} <- {len(grouped[job_name])} levels")
            continue
        out_path.write_text(merged, encoding="utf-8")
        created += 1
        for level_path in grouped[job_name].values():
            if level_path != out_path:
                level_path.unlink()
                removed += 1

    if args.dry_run:
        print(f"将生成 {len(grouped)} 个合并文件，删除 {sum(len(v) for v in grouped.values())} 个职级文件")
        return 0

    remaining = list(jobmap_dir.glob("*-*.md"))
    print(f"已生成 {created} 个合并文件，删除 {removed} 个职级文件，剩余 {len(list(jobmap_dir.glob('*.md')))} 个 .md")
    if remaining:
        print(f"警告：仍有 {len(remaining)} 个带连字符的文件未处理")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
