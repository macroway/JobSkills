#!/usr/bin/env python3
"""Build a one-page Chinese resume by cloning the bundled DOCX template.

The script is intentionally content-agnostic. It accepts structured JSON,
reuses paragraph/run prototypes from the retained template, and preserves
untouched package parts. Optional local font embedding is supported without
copying any font file into the skill directory.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import struct
import sys
from copy import deepcopy
from pathlib import Path
from uuid import NAMESPACE_URL, uuid5
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "中文简历模版.docx"
EXPECTED_TEMPLATE_SHA256 = (
    "252897222cfedf3973abd090f057779b22483d3dfa616410ad41f87efdbe2fbf"
)
DEFAULT_FONT = "方正楷体_GB2312"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W}


class ResumeBuildError(ValueError):
    """Raised when source data or the template cannot be used safely."""


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


def run_rpr(run: etree._Element) -> etree._Element:
    rpr = run.find(qn("rPr"))
    return deepcopy(rpr) if rpr is not None else etree.Element(qn("rPr"))


def is_bold(rpr: etree._Element | None) -> bool:
    if rpr is None:
        return False
    bold = rpr.find(qn("b"))
    if bold is None:
        return False
    return bold.get(qn("val"), "1") not in {"0", "false", "off"}


def set_font(rpr: etree._Element, font_name: str) -> None:
    fonts = rpr.find(qn("rFonts"))
    if fonts is None:
        fonts = etree.Element(qn("rFonts"))
        rpr.insert(0, fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(attr), font_name)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(attr), None)


def set_bold(rpr: etree._Element, enabled: bool) -> None:
    for tag in ("b", "bCs"):
        for node in list(rpr.findall(qn(tag))):
            rpr.remove(node)
    if enabled:
        etree.SubElement(rpr, qn("b"))
        etree.SubElement(rpr, qn("bCs"))


def set_highlight(rpr: etree._Element, value: str | None) -> None:
    for node in list(rpr.findall(qn("highlight"))):
        rpr.remove(node)
    if value:
        highlight = etree.SubElement(rpr, qn("highlight"))
        highlight.set(qn("val"), value)


def clone_empty(paragraph: etree._Element, font_name: str) -> etree._Element:
    cloned = deepcopy(paragraph)
    for child in list(cloned):
        if child.tag != qn("pPr"):
            cloned.remove(child)
    paragraph_rpr = cloned.find("./w:pPr/w:rPr", namespaces=NS)
    if paragraph_rpr is not None:
        set_font(paragraph_rpr, font_name)
    return cloned


def add_run(
    paragraph: etree._Element,
    text: str,
    prototype_rpr: etree._Element,
    font_name: str,
    *,
    bold: bool | None = None,
    highlight: bool = False,
) -> etree._Element:
    run = etree.SubElement(paragraph, qn("r"))
    rpr = deepcopy(prototype_rpr)
    set_font(rpr, font_name)
    if bold is not None:
        set_bold(rpr, bold)
    set_highlight(rpr, "yellow" if highlight else None)
    run.append(rpr)
    text_node = etree.SubElement(run, qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(f"{{{XML}}}space", "preserve")
    text_node.text = text
    return run


def add_tab(
    paragraph: etree._Element,
    prototype_rpr: etree._Element,
    font_name: str,
) -> None:
    run = etree.SubElement(paragraph, qn("r"))
    rpr = deepcopy(prototype_rpr)
    set_font(rpr, font_name)
    run.append(rpr)
    etree.SubElement(run, qn("tab"))


def find_paragraph(
    paragraphs: list[etree._Element],
    predicate,
    description: str,
) -> etree._Element:
    for paragraph in paragraphs:
        if predicate(paragraph):
            return paragraph
    raise ResumeBuildError(f"模板中未找到必要槽位：{description}")


def first_text_run(paragraph: etree._Element) -> etree._Element:
    for run in paragraph.xpath("./w:r", namespaces=NS):
        if run_text(run):
            return run
    raise ResumeBuildError("模板段落缺少可复用的文字运行")


def extract_prototypes(root: etree._Element) -> dict[str, object]:
    body = root.find(qn("body"))
    if body is None:
        raise ResumeBuildError("模板缺少 word/document.xml 的 w:body")
    paragraphs = body.findall(qn("p"))

    name = find_paragraph(
        paragraphs,
        lambda p: style_id(p) == "6" and bool(paragraph_text(p)),
        "姓名",
    )
    contact = find_paragraph(
        paragraphs,
        lambda p: all(token in paragraph_text(p) for token in ("电话", "邮箱", "地点")),
        "联系方式",
    )
    heading = find_paragraph(
        paragraphs,
        lambda p: style_id(p) == "2" and paragraph_text(p).strip() == "教育背景",
        "一级栏目",
    )
    spacer = find_paragraph(
        paragraphs,
        lambda p: style_id(p) == "2" and not paragraph_text(p).strip(),
        "栏目间距",
    )
    entry = find_paragraph(
        paragraphs,
        lambda p: style_id(p) == "3" and bool(paragraph_text(p).strip()),
        "经历抬头",
    )
    bullet = find_paragraph(
        paragraphs,
        lambda p: style_id(p) == "12" and "政策拆析" in paragraph_text(p),
        "经历条目",
    )
    body_line = find_paragraph(
        paragraphs,
        lambda p: style_id(p) == "5" and "计算机能力" in paragraph_text(p),
        "个人技能",
    )
    photo = next(
        (
            p
            for p in paragraphs
            if p.xpath(".//w:drawing | .//w:pict", namespaces=NS)
        ),
        None,
    )

    bullet_runs = bullet.xpath("./w:r", namespaces=NS)
    marker_run = next(
        (
            run
            for run in bullet_runs
            if any(char in run_text(run) for char in ("\uf0b7", "•", "●"))
        ),
        None,
    )
    title_run = next(
        (
            run
            for run in bullet_runs
            if run is not marker_run
            and run_text(run).strip()
            and is_bold(run.find(qn("rPr")))
        ),
        None,
    )
    body_run = next(
        (
            run
            for run in bullet_runs
            if run is not marker_run
            and run_text(run).strip()
            and not is_bold(run.find(qn("rPr")))
        ),
        None,
    )
    if marker_run is None or title_run is None or body_run is None:
        raise ResumeBuildError("模板经历条目无法识别项目符号、短标题或正文格式")

    body_runs = [
        run
        for run in body_line.xpath("./w:r", namespaces=NS)
        if run_text(run).strip()
    ]
    label_run = next(
        (run for run in body_runs if is_bold(run.find(qn("rPr")))),
        body_runs[0] if body_runs else None,
    )
    text_run = next(
        (run for run in body_runs if not is_bold(run.find(qn("rPr")))),
        body_runs[-1] if body_runs else None,
    )
    if label_run is None or text_run is None:
        raise ResumeBuildError("模板个人技能段落无法识别标签和正文格式")

    return {
        "body": body,
        "sect_pr": deepcopy(body.find(qn("sectPr"))),
        "photo": photo,
        "name": name,
        "name_rpr": run_rpr(first_text_run(name)),
        "contact": contact,
        "contact_rpr": run_rpr(first_text_run(contact)),
        "heading": heading,
        "heading_rpr": run_rpr(first_text_run(heading)),
        "spacer": spacer,
        "entry": entry,
        "entry_rpr": run_rpr(first_text_run(entry)),
        "bullet": bullet,
        "marker_text": run_text(marker_run)[0],
        "marker_rpr": run_rpr(marker_run),
        "bullet_title_rpr": run_rpr(title_run),
        "bullet_body_rpr": run_rpr(body_run),
        "body_line": body_line,
        "body_label_rpr": run_rpr(label_run),
        "body_text_rpr": run_rpr(text_run),
    }


def require_string(value, field: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str):
        raise ResumeBuildError(f"{field} 必须是字符串")
    text = value.strip()
    if not allow_empty and not text:
        raise ResumeBuildError(f"{field} 不能为空")
    return text


def normalize_value(value, field: str) -> dict[str, object]:
    if isinstance(value, str):
        return {"text": require_string(value, field), "pending": False}
    if not isinstance(value, dict):
        raise ResumeBuildError(f"{field} 必须是字符串或含 text/pending 的对象")
    return {
        "text": require_string(value.get("text"), f"{field}.text"),
        "pending": bool(value.get("pending", False)),
    }


def normalize_spans(value, field: str) -> list[dict[str, object]]:
    if isinstance(value, str):
        return [{"text": require_string(value, field), "highlight": False}]
    if not isinstance(value, list) or not value:
        raise ResumeBuildError(f"{field} 必须是非空字符串或文本片段数组")
    spans: list[dict[str, object]] = []
    for index, span in enumerate(value):
        if not isinstance(span, dict):
            raise ResumeBuildError(f"{field}[{index}] 必须是对象")
        text = require_string(span.get("text"), f"{field}[{index}].text")
        highlight = bool(span.get("highlight", False))
        if highlight and not re.fullmatch(r"【请补充：[^】]+】", text):
            raise ResumeBuildError(
                f"{field}[{index}] 的黄色提示必须完整写成【请补充：具体信息】"
            )
        spans.append({"text": text, "highlight": highlight})
    return spans


def normalize_bullet(item, field: str) -> dict[str, object]:
    if not isinstance(item, dict):
        raise ResumeBuildError(f"{field} 必须是对象")
    title = require_string(item.get("title"), f"{field}.title")
    if not 4 <= len(title) <= 10:
        raise ResumeBuildError(f"{field}.title 应为 4–10 个字")
    return {
        "title": title,
        "body": normalize_spans(item.get("body"), f"{field}.body"),
    }


def normalize_data(raw) -> dict[str, object]:
    if not isinstance(raw, dict):
        raise ResumeBuildError("JSON 根节点必须是对象")
    name = require_string(raw.get("name"), "name")
    contact_raw = raw.get("contact")
    if not isinstance(contact_raw, dict):
        raise ResumeBuildError("contact 必须是对象")
    contact = {
        key: normalize_value(contact_raw.get(key), f"contact.{key}")
        for key in ("phone", "email", "location")
    }

    sections_raw = raw.get("sections")
    if not isinstance(sections_raw, list) or not sections_raw:
        raise ResumeBuildError("sections 必须是非空数组")
    sections: list[dict[str, object]] = []
    for section_index, section in enumerate(sections_raw):
        field = f"sections[{section_index}]"
        if not isinstance(section, dict):
            raise ResumeBuildError(f"{field} 必须是对象")
        title = require_string(section.get("title"), f"{field}.title")
        kind = section.get("kind")
        if kind not in {"bullets", "entries", "body"}:
            raise ResumeBuildError(f"{field}.kind 必须是 bullets、entries 或 body")
        items_raw = section.get("items")
        if not isinstance(items_raw, list) or not items_raw:
            raise ResumeBuildError(f"{field}.items 必须是非空数组")

        items: list[dict[str, object]] = []
        if kind == "bullets":
            items = [
                normalize_bullet(item, f"{field}.items[{index}]")
                for index, item in enumerate(items_raw)
            ]
        elif kind == "entries":
            for item_index, item in enumerate(items_raw):
                item_field = f"{field}.items[{item_index}]"
                if not isinstance(item, dict):
                    raise ResumeBuildError(f"{item_field} 必须是对象")
                bullets_raw = item.get("bullets", [])
                if not isinstance(bullets_raw, list):
                    raise ResumeBuildError(f"{item_field}.bullets 必须是数组")
                items.append(
                    {
                        "left": require_string(item.get("left"), f"{item_field}.left"),
                        "middle": require_string(
                            item.get("middle"), f"{item_field}.middle"
                        ),
                        "right": normalize_value(
                            item.get("right"), f"{item_field}.right"
                        ),
                        "bullets": [
                            normalize_bullet(
                                bullet, f"{item_field}.bullets[{bullet_index}]"
                            )
                            for bullet_index, bullet in enumerate(bullets_raw)
                        ],
                    }
                )
        else:
            for item_index, item in enumerate(items_raw):
                item_field = f"{field}.items[{item_index}]"
                if not isinstance(item, dict):
                    raise ResumeBuildError(f"{item_field} 必须是对象")
                items.append(
                    {
                        "label": require_string(
                            item.get("label"), f"{item_field}.label"
                        ),
                        "body": normalize_spans(
                            item.get("body"), f"{item_field}.body"
                        ),
                    }
                )
        sections.append({"title": title, "kind": kind, "items": items})

    photo_path = raw.get("photo_path")
    if photo_path is not None:
        photo_path = Path(require_string(photo_path, "photo_path")).expanduser()
        if not photo_path.is_file():
            raise ResumeBuildError(f"photo_path 不存在：{photo_path}")

    return {
        "name": name,
        "contact": contact,
        "sections": sections,
        "photo_path": photo_path,
    }


def make_title(proto, rpr, text: str, font_name: str) -> etree._Element:
    paragraph = clone_empty(proto, font_name)
    ppr = paragraph.find(qn("pPr"))
    if ppr is None:
        ppr = etree.SubElement(paragraph, qn("pPr"))
    indent = ppr.find(qn("ind"))
    if indent is None:
        indent = etree.SubElement(ppr, qn("ind"))
    indent.set(qn("left"), "4300")
    indent.set(qn("right"), "3821")
    add_run(paragraph, text, rpr, font_name, bold=True)
    return paragraph


def make_contact(proto, rpr, contact, font_name: str) -> etree._Element:
    paragraph = clone_empty(proto, font_name)
    values = [
        ("电话：", contact["phone"]),
        ("邮箱：", contact["email"]),
        ("地点：", contact["location"]),
    ]
    for index, (label, value) in enumerate(values):
        add_run(paragraph, label, rpr, font_name, bold=True)
        add_run(
            paragraph,
            value["text"],
            rpr,
            font_name,
            bold=True,
            highlight=bool(value["pending"]),
        )
        if index == 0:
            add_tab(paragraph, rpr, font_name)
        elif index == 1:
            add_tab(paragraph, rpr, font_name)
            add_tab(paragraph, rpr, font_name)
    return paragraph


def make_heading(proto, rpr, text: str, font_name: str) -> etree._Element:
    paragraph = clone_empty(proto, font_name)
    add_run(paragraph, text, rpr, font_name, bold=True)
    add_tab(paragraph, rpr, font_name)
    add_tab(paragraph, rpr, font_name)
    return paragraph


def make_entry(proto, rpr, item, font_name: str) -> etree._Element:
    paragraph = clone_empty(proto, font_name)
    add_run(paragraph, item["left"], rpr, font_name, bold=True)
    add_tab(paragraph, rpr, font_name)
    add_run(paragraph, item["middle"], rpr, font_name, bold=True)
    add_tab(paragraph, rpr, font_name)
    add_run(
        paragraph,
        item["right"]["text"],
        rpr,
        font_name,
        bold=True,
        highlight=bool(item["right"]["pending"]),
    )
    return paragraph


def make_bullet(prototypes, item, font_name: str) -> etree._Element:
    paragraph = clone_empty(prototypes["bullet"], font_name)
    add_run(
        paragraph,
        prototypes["marker_text"],
        prototypes["marker_rpr"],
        "Symbol",
    )
    add_run(
        paragraph,
        f"【★{item['title']}】",
        prototypes["bullet_title_rpr"],
        font_name,
        bold=True,
    )
    for span in item["body"]:
        add_run(
            paragraph,
            span["text"],
            prototypes["bullet_body_rpr"],
            font_name,
            bold=False,
            highlight=bool(span["highlight"]),
        )
    return paragraph


def make_body(prototypes, item, font_name: str) -> etree._Element:
    paragraph = clone_empty(prototypes["body_line"], font_name)
    add_run(
        paragraph,
        item["label"],
        prototypes["body_label_rpr"],
        font_name,
        bold=True,
    )
    for span in item["body"]:
        add_run(
            paragraph,
            span["text"],
            prototypes["body_text_rpr"],
            font_name,
            bold=False,
            highlight=bool(span["highlight"]),
        )
    return paragraph


def render_body(root, data, prototypes, font_name: str) -> None:
    body = prototypes["body"]
    for child in list(body):
        body.remove(child)

    if data["photo_path"] is not None:
        if prototypes["photo"] is None:
            raise ResumeBuildError("模板没有可替换的照片槽")
        body.append(deepcopy(prototypes["photo"]))
    body.append(make_title(
        prototypes["name"], prototypes["name_rpr"], data["name"], font_name
    ))
    body.append(make_contact(
        prototypes["contact"],
        prototypes["contact_rpr"],
        data["contact"],
        font_name,
    ))

    for section_index, section in enumerate(data["sections"]):
        if section_index:
            body.append(clone_empty(prototypes["spacer"], font_name))
        body.append(make_heading(
            prototypes["heading"],
            prototypes["heading_rpr"],
            section["title"],
            font_name,
        ))
        if section["kind"] == "bullets":
            for item in section["items"]:
                body.append(make_bullet(prototypes, item, font_name))
        elif section["kind"] == "entries":
            for item in section["items"]:
                body.append(make_entry(
                    prototypes["entry"],
                    prototypes["entry_rpr"],
                    item,
                    font_name,
                ))
                for bullet in item["bullets"]:
                    body.append(make_bullet(prototypes, bullet, font_name))
        else:
            for item in section["items"]:
                body.append(make_body(prototypes, item, font_name))

    if prototypes["sect_pr"] is None:
        raise ResumeBuildError("模板缺少节属性 w:sectPr")
    body.append(prototypes["sect_pr"])


def checksum32(data: bytes | bytearray) -> int:
    padding = (-len(data)) % 4
    if padding:
        data = bytes(data) + (b"\0" * padding)
    return sum(struct.unpack(f">{len(data) // 4}I", data)) & 0xFFFFFFFF


def extract_ttc_face(data: bytes, face_index: int) -> bytes:
    if data[:4] != b"ttcf":
        if face_index != 0:
            raise ResumeBuildError("非 TTC 字体的 --font-index 只能为 0")
        return data
    if len(data) < 12:
        raise ResumeBuildError("TTC 字体头损坏")
    num_fonts = struct.unpack_from(">I", data, 8)[0]
    if face_index < 0 or face_index >= num_fonts:
        raise ResumeBuildError(
            f"--font-index {face_index} 超出 TTC 字体范围 0–{num_fonts - 1}"
        )
    face_offset = struct.unpack_from(">I", data, 12 + 4 * face_index)[0]
    if face_offset + 12 > len(data):
        raise ResumeBuildError("TTC 字体中的面偏移无效")

    sfnt_version = data[face_offset : face_offset + 4]
    num_tables, search_range, entry_selector, range_shift = struct.unpack_from(
        ">HHHH", data, face_offset + 4
    )
    records = []
    for index in range(num_tables):
        record_offset = face_offset + 12 + index * 16
        tag, _old_checksum, source_offset, length = struct.unpack_from(
            ">4sIII", data, record_offset
        )
        if source_offset + length > len(data):
            raise ResumeBuildError(f"TTC 字体表 {tag!r} 超出文件范围")
        records.append((tag, data[source_offset : source_offset + length]))

    directory_size = 12 + 16 * num_tables
    cursor = (directory_size + 3) & ~3
    output = bytearray(cursor)
    output[:4] = sfnt_version
    struct.pack_into(
        ">HHHH",
        output,
        4,
        num_tables,
        search_range,
        entry_selector,
        range_shift,
    )

    head_output_offset = None
    for index, (tag, table_data) in enumerate(records):
        table = bytearray(table_data)
        if tag == b"head" and len(table) >= 12:
            table[8:12] = b"\0\0\0\0"
            head_output_offset = cursor
        table_checksum = checksum32(table)
        struct.pack_into(
            ">4sIII",
            output,
            12 + index * 16,
            tag,
            table_checksum,
            cursor,
            len(table),
        )
        output.extend(table)
        output.extend(b"\0" * ((-len(table)) % 4))
        cursor += len(table) + ((-len(table)) % 4)

    if head_output_offset is not None:
        adjustment = (0xB1B0AFBA - checksum32(output)) & 0xFFFFFFFF
        struct.pack_into(">I", output, head_output_offset + 8, adjustment)
    return bytes(output)


def font_embedding_allowed(font_data: bytes) -> None:
    if len(font_data) < 12:
        raise ResumeBuildError("字体文件过短")
    num_tables = struct.unpack_from(">H", font_data, 4)[0]
    for index in range(num_tables):
        offset = 12 + index * 16
        if offset + 16 > len(font_data):
            break
        tag, _checksum, table_offset, length = struct.unpack_from(
            ">4sIII", font_data, offset
        )
        if tag == b"OS/2" and length >= 10 and table_offset + 10 <= len(font_data):
            fs_type = struct.unpack_from(">H", font_data, table_offset + 8)[0]
            if fs_type & 0x0002:
                raise ResumeBuildError("字体授权标记禁止嵌入，不能写入 DOCX")
            if fs_type & 0x0100:
                raise ResumeBuildError("字体授权标记仅允许位图嵌入，不能写入 DOCX")
            return


def obfuscate_openxml_font(font_data: bytes, font_key: str) -> bytes:
    from uuid import UUID

    data = bytearray(font_data)
    key = UUID(font_key).bytes[::-1]
    for index in range(min(32, len(data))):
        data[index] ^= key[index % 16]
    return bytes(data)


def next_relationship_id(root: etree._Element) -> str:
    used = {node.get("Id") for node in root}
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def ensure_font_entry(font_table: etree._Element, font_name: str) -> etree._Element:
    for entry in font_table.findall(qn("font")):
        if entry.get(qn("name")) == font_name:
            return entry
    entry = etree.SubElement(font_table, qn("font"))
    entry.set(qn("name"), font_name)
    charset = etree.SubElement(entry, qn("charset"))
    charset.set(qn("val"), "86")
    family = etree.SubElement(entry, qn("family"))
    family.set(qn("val"), "modern")
    pitch = etree.SubElement(entry, qn("pitch"))
    pitch.set(qn("val"), "default")
    return entry


def prepare_named_font_parts(source_zip, font_name: str) -> dict[str, bytes]:
    """Switch document styles to a locally available font without embedding it."""
    font_table = etree.fromstring(source_zip.read("word/fontTable.xml"))
    for embedded in font_table.xpath(".//w:embedRegular", namespaces=NS):
        embedded.getparent().remove(embedded)
    ensure_font_entry(font_table, font_name)

    styles = etree.fromstring(source_zip.read("word/styles.xml"))
    for fonts in styles.xpath(".//w:rFonts", namespaces=NS):
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(attr), font_name)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn(attr), None)

    xml_options = {
        "xml_declaration": True,
        "encoding": "UTF-8",
        "standalone": True,
    }
    return {
        "word/fontTable.xml": etree.tostring(font_table, **xml_options),
        "word/styles.xml": etree.tostring(styles, **xml_options),
    }


def prepare_font_parts(source_zip, font_path: Path, font_index: int, font_name: str):
    raw_font = font_path.read_bytes()
    standalone_font = extract_ttc_face(raw_font, font_index)
    font_embedding_allowed(standalone_font)
    fingerprint = hashlib.sha256(standalone_font).hexdigest()
    font_key = str(uuid5(NAMESPACE_URL, f"{font_name}:{fingerprint}")).upper()

    font_table = etree.fromstring(source_zip.read("word/fontTable.xml"))
    for embedded in font_table.xpath(".//w:embedRegular", namespaces=NS):
        embedded.getparent().remove(embedded)
    entry = ensure_font_entry(font_table, font_name)

    relationships = etree.fromstring(
        source_zip.read("word/_rels/fontTable.xml.rels")
    )
    relationship_id = next_relationship_id(relationships)
    embedded = etree.SubElement(entry, qn("embedRegular"))
    embedded.set(f"{{{R}}}id", relationship_id)
    embedded.set(qn("fontKey"), f"{{{font_key}}}")

    relationship = etree.SubElement(
        relationships, f"{{{PKG_REL}}}Relationship"
    )
    relationship.set("Id", relationship_id)
    relationship.set(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font",
    )
    relationship.set("Target", "fonts/resume-kaiti.odttf")

    styles = etree.fromstring(source_zip.read("word/styles.xml"))
    for fonts in styles.xpath(".//w:rFonts", namespaces=NS):
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(attr), font_name)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn(attr), None)

    content_types = etree.fromstring(source_zip.read("[Content_Types].xml"))
    part_name = "/word/fonts/resume-kaiti.odttf"
    if not any(
        node.get("PartName") == part_name
        for node in content_types.findall(f"{{{CONTENT_TYPES}}}Override")
    ):
        override = etree.SubElement(
            content_types, f"{{{CONTENT_TYPES}}}Override"
        )
        override.set("PartName", part_name)
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.obfuscatedFont",
        )

    xml_options = {
        "xml_declaration": True,
        "encoding": "UTF-8",
        "standalone": True,
    }
    replacements = {
        "word/fontTable.xml": etree.tostring(font_table, **xml_options),
        "word/_rels/fontTable.xml.rels": etree.tostring(
            relationships, **xml_options
        ),
        "word/styles.xml": etree.tostring(styles, **xml_options),
        "[Content_Types].xml": etree.tostring(content_types, **xml_options),
    }
    additions = {
        "word/fonts/resume-kaiti.odttf": obfuscate_openxml_font(
            standalone_font, font_key
        )
    }
    return replacements, additions


def prepare_photo(source_zip, photo_path: Path) -> bytes:
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise ResumeBuildError("替换照片需要 Pillow") from exc
    placeholder = Image.open(io.BytesIO(source_zip.read("word/media/image1.png")))
    width, height = placeholder.size
    with Image.open(photo_path) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
        image = ImageOps.fit(
            image,
            (width, height),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.35),
        )
        output = io.BytesIO()
        image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def build(
    template: Path,
    output: Path,
    data_path: Path,
    font_name: str,
    font_path: Path | None,
    font_index: int,
) -> None:
    if not template.is_file():
        raise ResumeBuildError(f"模板不存在：{template}")
    if sha256_file(template) != EXPECTED_TEMPLATE_SHA256:
        raise ResumeBuildError(
            "模板 SHA-256 与交付标准不一致；请恢复内置模板后再生成"
        )
    if output.resolve() == template.resolve():
        raise ResumeBuildError("输出路径不能覆盖模板原件")
    if font_path is not None and not font_path.is_file():
        raise ResumeBuildError(f"字体文件不存在：{font_path}")

    with data_path.open("r", encoding="utf-8") as stream:
        data = normalize_data(json.load(stream))

    xml_options = {
        "xml_declaration": True,
        "encoding": "UTF-8",
        "standalone": True,
    }
    with ZipFile(template) as source_zip:
        document_root = etree.fromstring(source_zip.read("word/document.xml"))
        prototypes = extract_prototypes(document_root)
        render_body(document_root, data, prototypes, font_name)
        replacements = {
            "word/document.xml": etree.tostring(document_root, **xml_options)
        }
        additions: dict[str, bytes] = {}

        if font_path is not None:
            font_replacements, font_additions = prepare_font_parts(
                source_zip, font_path, font_index, font_name
            )
            replacements.update(font_replacements)
            additions.update(font_additions)
        elif font_name != DEFAULT_FONT:
            replacements.update(prepare_named_font_parts(source_zip, font_name))

        if data["photo_path"] is not None:
            replacements["word/media/image1.png"] = prepare_photo(
                source_zip, data["photo_path"]
            )

        output.parent.mkdir(parents=True, exist_ok=True)
        temporary = output.with_name(f".{output.name}.tmp")
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as target_zip:
            for info in source_zip.infolist():
                target_zip.writestr(
                    info,
                    replacements.get(info.filename, source_zip.read(info.filename)),
                )
            for part_name, part_data in additions.items():
                if part_name not in source_zip.namelist():
                    target_zip.writestr(part_name, part_data)
        temporary.replace(output)

    if sha256_file(template) != EXPECTED_TEMPLATE_SHA256:
        raise ResumeBuildError("生成后模板哈希发生变化，已停止交付")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="基于内置中文简历模板和结构化 JSON 生成 Word 简历"
    )
    parser.add_argument("--data", type=Path, required=True, help="UTF-8 JSON 数据")
    parser.add_argument("--output", type=Path, required=True, help="输出 DOCX")
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help="模板 DOCX；默认使用技能内置模板",
    )
    parser.add_argument(
        "--font-name",
        default=DEFAULT_FONT,
        help=f"中文字体名称；默认 {DEFAULT_FONT}",
    )
    parser.add_argument(
        "--font-file",
        type=Path,
        help="可合法嵌入的本地 TTF/OTF/TTC 字体；不会复制到技能目录",
    )
    parser.add_argument(
        "--font-index",
        type=int,
        default=0,
        help="TTC 字体面序号；TTF/OTF 必须为 0",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        build(
            args.template.resolve(),
            args.output.resolve(),
            args.data.resolve(),
            require_string(args.font_name, "--font-name"),
            args.font_file.resolve() if args.font_file else None,
            args.font_index,
        )
    except (ResumeBuildError, json.JSONDecodeError, OSError, etree.XMLSyntaxError) as exc:
        print(f"[build_resume_docx] ERROR: {exc}", file=sys.stderr)
        return 2
    print(args.output.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
