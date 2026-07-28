#!/usr/bin/env python3
"""Audit structural delivery requirements for a generated resume DOCX."""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree


SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "中文简历模版.docx"
EXPECTED_TEMPLATE_SHA256 = (
    "252897222cfedf3973abd090f057779b22483d3dfa616410ad41f87efdbe2fbf"
)
DEFAULT_FONT = "方正楷体_GB2312"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W}

PLACEHOLDER_RE = re.compile(r"【请补充：[^】]+】")
TITLE_RE = re.compile(r"^[\uf0b7•●·\s]*【★([^】]+)】")
TEMPLATE_LEFTOVERS = (
    "XXXXX",
    "XXXXX大学",
    "JOBLOGIC-X",
    "第一二点结合岗位需求",
    "第三点可以体现",
    "核心写工作技能",
    "着重写软能力",
    "科研背景,突出软技能",
    "相片",
)


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def style_id(paragraph: etree._Element) -> str | None:
    values = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return values[0] if values else None


def run_text(run: etree._Element) -> str:
    return "".join(run.xpath(".//w:t/text()", namespaces=NS))


def geometry_signature(root: etree._Element) -> list[tuple]:
    signatures = []
    for section in root.xpath("//w:sectPr", namespaces=NS):
        page_size = section.find(qn("pgSz"))
        margins = section.find(qn("pgMar"))
        signatures.append(
            (
                tuple(
                    page_size.get(qn(attr))
                    for attr in ("w", "h", "orient", "code")
                )
                if page_size is not None
                else None,
                tuple(
                    margins.get(qn(attr))
                    for attr in (
                        "top",
                        "right",
                        "bottom",
                        "left",
                        "header",
                        "footer",
                        "gutter",
                    )
                )
                if margins is not None
                else None,
            )
        )
    return signatures


def highlight_value(run: etree._Element) -> str | None:
    values = run.xpath("./w:rPr/w:highlight/@w:val", namespaces=NS)
    return values[0] if values else None


def audit(args: argparse.Namespace) -> tuple[list[str], list[str], dict[str, int]]:
    failures: list[str] = []
    warnings: list[str] = []
    metrics = {"yellow": 0, "ability_titles": 0, "pages": 0}

    if not args.docx.is_file():
        return [f"文件不存在：{args.docx}"], warnings, metrics
    if not args.template.is_file():
        return [f"模板不存在：{args.template}"], warnings, metrics
    template_hash = sha256_file(args.template)
    if template_hash != EXPECTED_TEMPLATE_SHA256:
        failures.append(
            f"模板 SHA-256 不一致：{template_hash}"
        )

    try:
        with ZipFile(args.docx) as resume_zip:
            broken_part = resume_zip.testzip()
            if broken_part:
                failures.append(f"DOCX 压缩包损坏：{broken_part}")
            required = {
                "[Content_Types].xml",
                "word/document.xml",
                "word/styles.xml",
                "word/fontTable.xml",
            }
            missing = sorted(required.difference(resume_zip.namelist()))
            if missing:
                failures.append(f"DOCX 缺少必要部件：{', '.join(missing)}")
                return failures, warnings, metrics
            root = etree.fromstring(resume_zip.read("word/document.xml"))
            font_table = etree.fromstring(resume_zip.read("word/fontTable.xml"))
    except (BadZipFile, OSError, etree.XMLSyntaxError) as exc:
        return [f"DOCX 无法读取：{exc}"], warnings, metrics

    with ZipFile(args.template) as template_zip:
        template_root = etree.fromstring(template_zip.read("word/document.xml"))

    if geometry_signature(root) != geometry_signature(template_root):
        failures.append("页面尺寸、页边距或节结构与模板不一致")

    if root.xpath("//w:ins | //w:del", namespaces=NS):
        failures.append("DOCX 仍包含未接受的修订")

    document_text = "\n".join(
        paragraph_text(paragraph)
        for paragraph in root.xpath("//w:body/w:p", namespaces=NS)
    )
    for leftover in TEMPLATE_LEFTOVERS:
        if leftover in document_text:
            failures.append(f"残留模板示例或说明文字：{leftover}")
    if re.search(r"(?<![A-Za-z])X{2,}(?![A-Za-z])", document_text):
        failures.append("残留由两个以上 X 组成的模板占位符")

    for paragraph_index, paragraph in enumerate(
        root.xpath("//w:body/w:p", namespaces=NS), start=1
    ):
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        if style_id(paragraph) == "12" or "【★" in text:
            match = TITLE_RE.match(text)
            if match is None:
                failures.append(
                    f"第 {paragraph_index} 段经历条目未以【★能力短标题】开头"
                )
            else:
                title = match.group(1)
                metrics["ability_titles"] += 1
                if not 4 <= len(title) <= 10:
                    failures.append(
                        f"第 {paragraph_index} 段能力短标题不是 4–10 个字：{title}"
                    )

    for run in root.xpath("//w:r", namespaces=NS):
        text = run_text(run)
        if not text:
            continue
        highlight = highlight_value(run)
        if highlight == "yellow":
            metrics["yellow"] += 1
            if PLACEHOLDER_RE.fullmatch(text) is None:
                failures.append(f"黄色高亮不是完整待补提示：{text}")
        for placeholder in PLACEHOLDER_RE.findall(text):
            if highlight != "yellow" or text != placeholder:
                failures.append(f"待补提示未被完整黄色高亮：{placeholder}")

        if all(character in {"\uf0b7", "•", "●", "·", " "} for character in text):
            continue
        fonts = run.find("./w:rPr/w:rFonts", namespaces=NS)
        if fonts is None:
            failures.append(f"文字运行缺少字体属性：{text[:20]}")
            continue
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            if fonts.get(qn(attr)) != args.font_name:
                failures.append(
                    f"字体不一致：{text[:20]} 的 {attr}="
                    f"{fonts.get(qn(attr))!r}，应为 {args.font_name!r}"
                )
                break

    if args.require_embedded_font:
        matching_fonts = [
            entry
            for entry in font_table.findall(qn("font"))
            if entry.get(qn("name")) == args.font_name
        ]
        if not matching_fonts or not matching_fonts[0].xpath(
            "./w:embedRegular", namespaces=NS
        ):
            failures.append(f"字体 {args.font_name} 未嵌入 DOCX")

    if args.render_dir is not None:
        pages = sorted(args.render_dir.glob("page-*.png"))
        metrics["pages"] = len(pages)
        if len(pages) != args.expected_pages:
            failures.append(
                f"渲染页数为 {len(pages)}，应为 {args.expected_pages}"
            )
    else:
        warnings.append("未提供 --render-dir，无法验证实际渲染页数")

    try:
        from docx import Document

        Document(str(args.docx))
    except ImportError:
        warnings.append("当前环境没有 python-docx，跳过 Word 包级打开测试")
    except Exception as exc:  # python-docx reports package errors with varied types.
        failures.append(f"python-docx 无法打开文件：{exc}")

    return failures, warnings, metrics


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="检查中文简历 DOCX 的模板、字体、黄标和一页交付要求"
    )
    parser.add_argument("docx", type=Path, help="待检查 DOCX")
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help="模板 DOCX；默认使用技能内置模板",
    )
    parser.add_argument(
        "--font-name",
        default=DEFAULT_FONT,
        help=f"预期中文字体；默认 {DEFAULT_FONT}",
    )
    parser.add_argument(
        "--require-embedded-font",
        action="store_true",
        help="要求预期字体在 fontTable 中具有嵌入关系",
    )
    parser.add_argument(
        "--render-dir",
        type=Path,
        help="标准渲染产生 page-*.png 的目录",
    )
    parser.add_argument(
        "--expected-pages",
        type=int,
        default=1,
        help="预期页数；默认 1",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    args.docx = args.docx.resolve()
    args.template = args.template.resolve()
    if args.render_dir is not None:
        args.render_dir = args.render_dir.resolve()
    failures, warnings, metrics = audit(args)

    for warning in warnings:
        print(f"[audit_resume_docx] WARNING: {warning}")
    if failures:
        for failure in failures:
            print(f"[audit_resume_docx] FAIL: {failure}", file=sys.stderr)
        print(
            "[audit_resume_docx] RESULT: FAIL "
            f"(能力标题 {metrics['ability_titles']}，黄标 {metrics['yellow']}，"
            f"页数 {metrics['pages'] or '未验证'})",
            file=sys.stderr,
        )
        return 1
    print(
        "[audit_resume_docx] RESULT: PASS "
        f"(能力标题 {metrics['ability_titles']}，黄标 {metrics['yellow']}，"
        f"页数 {metrics['pages'] or '未验证'})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
