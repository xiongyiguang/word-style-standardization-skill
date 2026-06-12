#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 标准样式规整脚本

处理原则：
- 不重建正文，不抽取文本后重新生成 docx。
- 复制源 docx 包结构，仅替换/补充模板样式文件，并有限修改段落样式。
- 尽量保留图片、媒体、页眉页脚、对象锚点、关系引用和嵌入对象。
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W, "a": A, "r": R}

STYLE = {
    "heading": {1: "10", 2: "20", 3: "30", 4: "40", 5: "50", 6: "6", 7: "7", 8: "8", 9: "9"},
    "body": "P1505",
    "bold_body": "P",
    "table": "P5",
    "image": "P6",
    "list1": "P2",
    "list2": "P3",
    "list3": "P4",
}

HEADING_STYLE_IDS = set(STYLE["heading"].values())
BODY_STYLE_IDS = {STYLE["body"], STYLE["bold_body"], STYLE["table"], STYLE["list1"], STYLE["list2"], STYLE["list3"]}
PARAGRAPH_DIRECT_FORMAT_TAGS = {"rPr", "spacing", "ind"}
RUN_DIRECT_FORMAT_TAGS = {"rFonts", "sz", "szCs", "color"}
HEADING_RUN_DIRECT_FORMAT_TAGS = RUN_DIRECT_FORMAT_TAGS | {"b", "bCs", "i", "iCs", "u", "highlight"}
TABLE_CELL_DIRECT_FORMAT_TAGS = {"tcBorders", "shd"}
TABLE_CELL_WIDTH_TAGS = {"tcW"}

# 候选 XML 文件：正文、页眉页脚、脚注尾注、批注、文本框等常见正文载体。
XML_CANDIDATE_PATTERNS = (
    re.compile(r"^word/document\.xml$"),
    re.compile(r"^word/header\d+\.xml$"),
    re.compile(r"^word/footer\d+\.xml$"),
    re.compile(r"^word/footnotes\.xml$"),
    re.compile(r"^word/endnotes\.xml$"),
    re.compile(r"^word/comments.*\.xml$"),
)

HEADING_PATTERNS = [
    # 1、1.1、1.1.1、1.1.1.1 等中文方案文档常见标题编号
    re.compile(r"^\s*(\d+(?:\.\d+){0,8})[\s　]+\S+"),
    # 第X章/节/部分
    re.compile(r"^\s*第[一二三四五六七八九十百千万0-9]+[章节部分篇][\s　、：:.-]*\S*"),
]

BULLET_MARKERS = [
    re.compile(r"^\s*[•●▪◦◆◇■□▶▷→✓✔※]\s*"),
    re.compile(r"^\s*[-–—]\s+"),
]


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    uri = {"w": W, "a": A, "r": R}[prefix]
    return f"{{{uri}}}{local}"


def unzip_docx(path: Path, target: Path) -> None:
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(target)


def zip_dir(source_dir: Path, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in source_dir.rglob("*"):
            if file.is_file():
                zf.write(file, file.relative_to(source_dir).as_posix())


def count_docx(path: Path) -> Dict[str, int]:
    result = {
        "media_files": 0,
        "image_relationships": 0,
        "paragraphs": 0,
        "tables": 0,
        "normal_style_paragraphs": 0,
        "numbered_paragraphs": 0,
    }
    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        result["media_files"] = len([n for n in names if n.startswith("word/media/") and not n.endswith("/")])
        for n in names:
            if n.endswith(".rels"):
                try:
                    text = zf.read(n).decode("utf-8", errors="ignore")
                except Exception:
                    continue
                result["image_relationships"] += text.count("/image")
        for n in names:
            if any(p.match(n) for p in XML_CANDIDATE_PATTERNS):
                try:
                    root = etree.fromstring(zf.read(n))
                except Exception:
                    continue
                result["paragraphs"] += len(root.xpath(".//w:p", namespaces=NS))
                result["tables"] += len(root.xpath(".//w:tbl", namespaces=NS))
                result["numbered_paragraphs"] += len(root.xpath(".//w:pPr/w:numPr", namespaces=NS))
                for p in root.xpath(".//w:p", namespaces=NS):
                    sid = get_pstyle(p)
                    if sid in (None, "Normal", "a"):
                        # 空段落、特殊域段落也会被计入，报告中只作为风险提示。
                        result["normal_style_paragraphs"] += 1
    return result


def copy_template_parts(template_dir: Path, work_dir: Path) -> List[str]:
    copied = []
    parts = [
        "word/styles.xml",
        "word/numbering.xml",
        "word/theme/theme1.xml",
        "word/fontTable.xml",
    ]
    for rel in parts:
        src = template_dir / rel
        dst = work_dir / rel
        if src.exists():
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            copied.append(rel)
    return copied


def parse_table_style_id(styles_file: Path, style_name: str = "表格标准样式") -> Optional[str]:
    if not styles_file.exists():
        return None

    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.parse(str(styles_file), parser).getroot()
    fallback = None
    for style in root.xpath(".//w:style[@w:type='table']", namespaces=NS):
        sid = style.get(qn("w:styleId"))
        name_node = style.find(qn("w:name"))
        name = (name_node.get(qn("w:val")) if name_node is not None else "").strip()
        if not sid:
            continue
        if name == style_name:
            return sid
        if style_name in name or name in style_name:
            fallback = fallback or sid
    return fallback


def parse_template_table_look(document_file: Path, table_style_id: Optional[str]) -> Optional[etree._Element]:
    if not table_style_id or not document_file.exists():
        return None

    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.parse(str(document_file), parser).getroot()
    for tbl in root.xpath(".//w:tbl", namespaces=NS):
        sid = tbl.xpath("./w:tblPr/w:tblStyle/@w:val", namespaces=NS)
        if sid and sid[0] == table_style_id:
            look = tbl.find("./w:tblPr/w:tblLook", namespaces=NS)
            if look is not None:
                return look
    return None


def int_to_chinese(num: int) -> str:
    digits = "零一二三四五六七八九"
    units = ["", "十", "百", "千"]
    if num <= 0:
        return str(num)
    if num < 10:
        return digits[num]
    if num < 20:
        return "十" + (digits[num % 10] if num % 10 else "")
    parts = []
    chars = list(str(num))
    length = len(chars)
    zero = False
    for idx, ch in enumerate(chars):
        value = int(ch)
        pos = length - idx - 1
        if value == 0:
            zero = bool(parts)
            continue
        if zero:
            parts.append("零")
            zero = False
        parts.append(digits[value] + (units[pos] if pos < len(units) else ""))
    return "".join(parts)


def format_number(value: int, fmt: str) -> str:
    if fmt in {"decimal", "decimalZero"}:
        return str(value).zfill(2) if fmt == "decimalZero" else str(value)
    if fmt in {"chineseCounting", "chineseCountingThousand", "chineseLegalSimplified", "ideographDigital"}:
        return int_to_chinese(value)
    if fmt == "upperLetter":
        return chr(ord("A") + value - 1) if 1 <= value <= 26 else str(value)
    if fmt == "lowerLetter":
        return chr(ord("a") + value - 1) if 1 <= value <= 26 else str(value)
    return str(value)


def parse_numbering(numbering_file: Path) -> Dict[str, Dict[int, Dict[str, str]]]:
    if not numbering_file.exists():
        return {}

    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.parse(str(numbering_file), parser).getroot()
    abstract_levels: Dict[str, Dict[int, Dict[str, str]]] = {}
    num_to_abstract: Dict[str, str] = {}

    for abstract in root.xpath("./w:abstractNum", namespaces=NS):
        aid = abstract.get(qn("w:abstractNumId"))
        if not aid:
            continue
        levels: Dict[int, Dict[str, str]] = {}
        for lvl in abstract.xpath("./w:lvl", namespaces=NS):
            ilvl = lvl.get(qn("w:ilvl"))
            if ilvl is None or not ilvl.isdigit():
                continue
            start_node = lvl.find(qn("w:start"))
            fmt_node = lvl.find(qn("w:numFmt"))
            text_node = lvl.find(qn("w:lvlText"))
            levels[int(ilvl)] = {
                "start": start_node.get(qn("w:val")) if start_node is not None else "1",
                "fmt": fmt_node.get(qn("w:val")) if fmt_node is not None else "decimal",
                "text": text_node.get(qn("w:val")) if text_node is not None else f"%{int(ilvl) + 1}.",
                "left": "",
            }
            ind_node = lvl.find("./w:pPr/w:ind", namespaces=NS)
            if ind_node is not None and ind_node.get(qn("w:left")):
                levels[int(ilvl)]["left"] = ind_node.get(qn("w:left"))
        abstract_levels[aid] = levels

    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("w:numId"))
        abstract_node = num.find(qn("w:abstractNumId"))
        if num_id and abstract_node is not None and abstract_node.get(qn("w:val")):
            num_to_abstract[num_id] = abstract_node.get(qn("w:val"))

    return {num_id: abstract_levels.get(aid, {}) for num_id, aid in num_to_abstract.items()}


class NumberingMaterializer:
    def __init__(self, definitions: Dict[str, Dict[int, Dict[str, str]]]):
        self.definitions = definitions
        self.counters: Dict[str, Dict[int, int]] = {}

    def level_info(self, p: etree._Element) -> Optional[Tuple[str, int, Dict[str, str]]]:
        nodes = p.xpath("./w:pPr/w:numPr", namespaces=NS)
        if not nodes:
            return None
        num_id_node = nodes[0].find(qn("w:numId"))
        ilvl_node = nodes[0].find(qn("w:ilvl"))
        if num_id_node is None:
            return None
        num_id = num_id_node.get(qn("w:val"))
        ilvl_text = ilvl_node.get(qn("w:val")) if ilvl_node is not None else "0"
        if not num_id or not ilvl_text or not ilvl_text.isdigit():
            return None
        ilvl = int(ilvl_text)
        levels = self.definitions.get(num_id)
        if not levels or ilvl not in levels:
            return None
        return num_id, ilvl, levels[ilvl]

    def is_bullet(self, p: etree._Element) -> bool:
        info = self.level_info(p)
        return bool(info and info[2].get("fmt") == "bullet")

    def bullet_level(self, p: etree._Element) -> Optional[int]:
        info = self.level_info(p)
        if not info or info[2].get("fmt") != "bullet":
            return None
        indent_level = list_level_from_left_indent(info[2].get("left"))
        if indent_level:
            return indent_level
        return min(info[1] + 1, 3)

    def text_for(self, p: etree._Element) -> Optional[str]:
        info = self.level_info(p)
        if not info:
            return None
        num_id, ilvl, level_def = info
        if level_def.get("fmt") == "bullet":
            return None
        levels = self.definitions.get(num_id)
        if not levels:
            return None

        counters = self.counters.setdefault(num_id, {})
        start = int(level_def.get("start", "1")) if level_def.get("start", "1").isdigit() else 1
        counters[ilvl] = counters.get(ilvl, start - 1) + 1
        for lower_level in list(counters):
            if lower_level > ilvl:
                del counters[lower_level]

        text = level_def.get("text", f"%{ilvl + 1}.")
        for idx in range(9):
            value = counters.get(idx)
            if value is None:
                continue
            fmt = levels.get(idx, {}).get("fmt", "decimal")
            text = text.replace(f"%{idx + 1}", format_number(value, fmt))
        return text.strip()


def insert_prefix_text(p: etree._Element, prefix: str) -> None:
    if not prefix:
        return
    current_text = get_text(p)
    if current_text.startswith(prefix):
        return
    run = etree.Element(qn("w:r"))
    text_node = etree.SubElement(run, qn("w:t"))
    text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = f"{prefix} "
    insert_at = 1 if p.find(qn("w:pPr")) is not None else 0
    p.insert(insert_at, run)


def parse_style_hints(styles_file: Path) -> Dict[str, object]:
    """读取源文档样式定义，避免替换模板样式后丢失原 styleId 语义。"""
    hints: Dict[str, object] = {
        "heading": {},
        "body": set(),
        "table": set(),
        "image": set(),
        "bold": set(),
        "list": {},
    }
    if not styles_file.exists():
        return hints

    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.parse(str(styles_file), parser).getroot()
    names: Dict[str, str] = {}
    based_on: Dict[str, str] = {}
    paragraph_styles: Set[str] = set()

    for style in root.xpath(".//w:style[@w:type='paragraph']", namespaces=NS):
        sid = style.get(qn("w:styleId"))
        if not sid:
            continue
        paragraph_styles.add(sid)
        name_node = style.find(qn("w:name"))
        based_node = style.find(qn("w:basedOn"))
        names[sid] = (name_node.get(qn("w:val")) if name_node is not None else "").strip()
        if based_node is not None and based_node.get(qn("w:val")):
            based_on[sid] = based_node.get(qn("w:val"))

    def direct_heading_level(sid: str) -> Optional[int]:
        name = names.get(sid, "")
        for value in (sid, name):
            m = re.search(r"(?:heading|标题|级)\s*([1-9])", value, re.I)
            if m:
                return min(max(int(m.group(1)), 1), 9)
        template_mapping = {"10": 1, "20": 2, "30": 3, "40": 4, "50": 5, "6": 6, "7": 7, "8": 8, "9": 9}
        return template_mapping.get(sid)

    heading_cache: Dict[str, Optional[int]] = {}

    def inherited_heading_level(sid: str, seen: Optional[Set[str]] = None) -> Optional[int]:
        if sid in heading_cache:
            return heading_cache[sid]
        if seen is None:
            seen = set()
        if sid in seen:
            return None
        seen.add(sid)
        level = direct_heading_level(sid)
        if level is None and sid in based_on:
            level = inherited_heading_level(based_on[sid], seen)
        heading_cache[sid] = level
        return level

    for sid in paragraph_styles:
        level = inherited_heading_level(sid)
        if level:
            hints["heading"][sid] = level
            continue

        name = names.get(sid, "")
        lower_name = name.lower()
        if any(token in name for token in ("表格", "表内", "表内容")) or "table" in lower_name:
            hints["table"].add(sid)
        elif any(token in name for token in ("图片", "图文", "图内容")) or "figure" in lower_name:
            hints["image"].add(sid)
        elif any(token in name for token in ("粗体", "重点")) or "bold" in lower_name:
            hints["bold"].add(sid)
        elif "箭头" in name or sid == STYLE["list1"]:
            hints["list"][sid] = 1
        elif "打钩" in name or sid == STYLE["list2"]:
            hints["list"][sid] = 2
        elif "四角星" in name or sid == STYLE["list3"]:
            hints["list"][sid] = 3
        elif any(token in name for token in ("正文", "段落", "body text")) or "normal" in lower_name:
            hints["body"].add(sid)

    return hints


def xml_files(work_dir: Path) -> Iterable[Path]:
    for file in (work_dir / "word").glob("*.xml"):
        rel = file.relative_to(work_dir).as_posix()
        if any(p.match(rel) for p in XML_CANDIDATE_PATTERNS):
            yield file


def get_text(p: etree._Element) -> str:
    texts = p.xpath(".//w:t/text()", namespaces=NS)
    return "".join(texts).strip()


def get_pstyle(p: etree._Element) -> Optional[str]:
    nodes = p.xpath("./w:pPr/w:pStyle", namespaces=NS)
    if not nodes:
        return None
    return nodes[0].get(qn("w:val"))


def ensure_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    return ppr


def ensure_tblpr(tbl: etree._Element) -> etree._Element:
    tblpr = tbl.find(qn("w:tblPr"))
    if tblpr is None:
        tblpr = etree.Element(qn("w:tblPr"))
        tbl.insert(0, tblpr)
    return tblpr


def set_pstyle(p: etree._Element, style_id: str) -> None:
    ppr = ensure_ppr(p)
    pstyle = ppr.find(qn("w:pStyle"))
    if pstyle is None:
        pstyle = etree.Element(qn("w:pStyle"))
        ppr.insert(0, pstyle)
    pstyle.set(qn("w:val"), style_id)


def set_table_style(tbl: etree._Element, style_id: str, table_look: Optional[etree._Element]) -> None:
    tblpr = ensure_tblpr(tbl)
    tblstyle = tblpr.find(qn("w:tblStyle"))
    if tblstyle is None:
        tblstyle = etree.Element(qn("w:tblStyle"))
        tblpr.insert(0, tblstyle)
    tblstyle.set(qn("w:val"), style_id)

    if table_look is not None:
        existing_look = tblpr.find(qn("w:tblLook"))
        if existing_look is not None:
            tblpr.remove(existing_look)
        tblpr.append(etree.fromstring(etree.tostring(table_look)))


def normalize_table_width(tbl: etree._Element) -> Counter:
    stats = Counter()
    tblpr = ensure_tblpr(tbl)
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = etree.Element(qn("w:tblW"))
        insert_at = 1 if tblpr.find(qn("w:tblStyle")) is not None else 0
        tblpr.insert(insert_at, tblw)
        stats["table_width_format"] += 1
    elif tblw.get(qn("w:w")) != "5000" or tblw.get(qn("w:type")) != "pct":
        stats["table_width_format"] += 1
    tblw.set(qn("w:w"), "5000")
    tblw.set(qn("w:type"), "pct")

    removed_layout = remove_children_by_local_name(tblpr, {"tblLayout"})
    if removed_layout:
        stats["table_width_format"] += removed_layout
    return stats


def remove_children_by_local_name(parent: Optional[etree._Element], local_names: Set[str]) -> int:
    if parent is None:
        return 0
    removed = 0
    for child in list(parent):
        if etree.QName(child).localname in local_names:
            parent.remove(child)
            removed += 1
    return removed


def cleanup_paragraph_direct_formatting(p: etree._Element, style_id: str) -> Counter:
    stats = Counter()
    ppr = p.find(qn("w:pPr"))
    if style_id in HEADING_STYLE_IDS or style_id in BODY_STYLE_IDS:
        removed = remove_children_by_local_name(ppr, PARAGRAPH_DIRECT_FORMAT_TAGS)
        if removed:
            stats["paragraph_direct_format"] += removed

    if style_id in HEADING_STYLE_IDS:
        run_tags = HEADING_RUN_DIRECT_FORMAT_TAGS
    elif style_id in BODY_STYLE_IDS:
        run_tags = RUN_DIRECT_FORMAT_TAGS
    else:
        run_tags = set()

    if run_tags:
        for rpr in p.xpath("./w:r/w:rPr", namespaces=NS):
            removed = remove_children_by_local_name(rpr, run_tags)
            if removed:
                stats["run_direct_format"] += removed
            if len(rpr) == 0 and not rpr.attrib:
                parent = rpr.getparent()
                if parent is not None:
                    parent.remove(rpr)
    return stats


def cleanup_table_cell_direct_formatting(tbl: etree._Element) -> Counter:
    stats = Counter()
    tbl_grid = tbl.find(qn("w:tblGrid"))
    if tbl_grid is not None:
        tbl.remove(tbl_grid)
        stats["table_column_width_format"] += 1

    for tcpr in tbl.xpath(".//w:tcPr", namespaces=NS):
        removed = remove_children_by_local_name(tcpr, TABLE_CELL_DIRECT_FORMAT_TAGS)
        if removed:
            stats["table_cell_direct_format"] += removed
        removed_width = remove_children_by_local_name(tcpr, TABLE_CELL_WIDTH_TAGS)
        if removed_width:
            stats["table_column_width_format"] += removed_width
    return stats


def remove_numpr(p: etree._Element) -> None:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return
    for node in ppr.findall(qn("w:numPr")):
        ppr.remove(node)


def has_image_or_object(p: etree._Element) -> bool:
    if p.xpath(".//w:drawing | .//w:pict | .//w:object | .//a:blip", namespaces=NS):
        return True
    return False


def is_in_table(p: etree._Element) -> bool:
    parent = p.getparent()
    while parent is not None:
        if parent.tag == qn("w:tc"):
            return True
        parent = parent.getparent()
    return False


def heading_level_from_existing(style_id: Optional[str], style_hints: Dict[str, object]) -> Optional[int]:
    if not style_id:
        return None
    heading_hints = style_hints.get("heading", {})
    if isinstance(heading_hints, dict) and style_id in heading_hints:
        return heading_hints[style_id]
    mapping = {"10": 1, "20": 2, "30": 3, "40": 4, "50": 5, "6": 6, "7": 7, "8": 8, "9": 9}
    if style_id in mapping:
        return mapping[style_id]
    m = re.match(r"heading\s*(\d+)", style_id, re.I)
    if m:
        return min(max(int(m.group(1)), 1), 9)
    return None


def heading_level_from_text(text: str) -> Optional[int]:
    if not text:
        return None
    m = HEADING_PATTERNS[0].match(text)
    if m:
        level = m.group(1).count(".") + 1
        return min(max(level, 1), 9)
    if HEADING_PATTERNS[1].match(text):
        return 1
    return None


def has_numpr(p: etree._Element) -> bool:
    return bool(p.xpath("./w:pPr/w:numPr", namespaces=NS))


def list_level_from_left_indent(left: Optional[str]) -> Optional[int]:
    if not left:
        return None
    try:
        left_twips = int(left)
    except ValueError:
        return None
    if left_twips <= 0:
        return None

    # The template/source docs use roughly 440 twips per visual bullet indent.
    # Some Word-generated values are 860 instead of 880, so use broad bands.
    if left_twips <= 660:
        return 1
    if left_twips <= 1100:
        return 2
    return 3


def paragraph_left_indent(p: etree._Element) -> Optional[str]:
    nodes = p.xpath("./w:pPr/w:ind", namespaces=NS)
    if not nodes:
        return None
    return nodes[0].get(qn("w:left"))


def bullet_level(p: etree._Element, text: str, numbering: NumberingMaterializer) -> Optional[int]:
    auto_bullet_level = numbering.bullet_level(p)
    if auto_bullet_level:
        return auto_bullet_level

    # 只按文本中的明确项目符号判断。自动编号和缩进只用于编号物化，
    # 不能触发 P2/P3/P4，否则会出现“箭头样式 + 1)”的重复编号。
    if any(pat.match(text) for pat in BULLET_MARKERS):
        indent_level = list_level_from_left_indent(paragraph_left_indent(p))
        if indent_level:
            return indent_level
        leading_spaces = len(text) - len(text.lstrip())
        if leading_spaces >= 4:
            return 3
        if leading_spaces >= 2:
            return 2
        return 1
    return None


def is_all_bold(p: etree._Element) -> bool:
    runs = p.xpath("./w:r", namespaces=NS)
    text_runs = [r for r in runs if r.xpath(".//w:t/text()", namespaces=NS)]
    if not text_runs:
        return False
    bold_runs = [r for r in text_runs if r.xpath("./w:rPr/w:b", namespaces=NS)]
    return len(bold_runs) == len(text_runs)


def style_hint_contains(style_hints: Dict[str, object], kind: str, style_id: Optional[str]) -> bool:
    values = style_hints.get(kind, set())
    return bool(style_id and isinstance(values, set) and style_id in values)


def list_level_from_style_hint(style_hints: Dict[str, object], style_id: Optional[str]) -> Optional[int]:
    values = style_hints.get("list", {})
    if not style_id or not isinstance(values, dict):
        return None
    level = values.get(style_id)
    if isinstance(level, int):
        return min(max(level, 1), 3)
    return None


def choose_style(p: etree._Element, style_hints: Dict[str, object], numbering: NumberingMaterializer) -> Tuple[str, bool]:
    """返回 (style_id, should_remove_numpr)。"""
    text = get_text(p)
    current = get_pstyle(p)

    if has_image_or_object(p) or style_hint_contains(style_hints, "image", current):
        return STYLE["image"], False

    # 标题只依据源文档已有样式语义判断。不要按文本编号猜标题，
    # 否则公式、金额、交易量等以数字开头的正文会被误套标题样式。
    level = heading_level_from_existing(current, style_hints)
    if level:
        return STYLE["heading"][level], False

    if is_in_table(p) or style_hint_contains(style_hints, "table", current):
        return STYLE["table"], True

    source_list_level = list_level_from_style_hint(style_hints, current)
    if source_list_level and (not has_numpr(p) or numbering.is_bullet(p)):
        return STYLE[f"list{source_list_level}"], True

    lvl = bullet_level(p, text, numbering)
    if lvl:
        return STYLE[f"list{lvl}"], True

    if text and (style_hint_contains(style_hints, "bold", current) or is_all_bold(p)) and len(text) <= 80:
        return STYLE["bold_body"], True

    return STYLE["body"], True


def normalize_xml(
    file: Path,
    style_hints: Dict[str, object],
    numbering: NumberingMaterializer,
    table_style_id: Optional[str],
    table_look: Optional[etree._Element],
) -> Tuple[Counter, Counter, Counter]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.parse(str(file), parser).getroot()
    stats = Counter()
    table_stats = Counter()
    cleanup_stats = Counter()

    if table_style_id:
        for tbl in root.xpath(".//w:tbl", namespaces=NS):
            set_table_style(tbl, table_style_id, table_look)
            cleanup_stats.update(normalize_table_width(tbl))
            cleanup_stats.update(cleanup_table_cell_direct_formatting(tbl))
            table_stats[table_style_id] += 1

    for p in root.xpath(".//w:p", namespaces=NS):
        style_id, remove_numbering = choose_style(p, style_hints, numbering)
        set_pstyle(p, style_id)
        cleanup_stats.update(cleanup_paragraph_direct_formatting(p, style_id))
        if remove_numbering:
            prefix = numbering.text_for(p)
            if prefix:
                insert_prefix_text(p, prefix)
            remove_numpr(p)
        stats[style_id] += 1

    tree = etree.ElementTree(root)
    tree.write(str(file), xml_declaration=True, encoding="UTF-8", standalone=True)
    return stats, table_stats, cleanup_stats


def generate_report(
    report: Path,
    input_file: Path,
    output_file: Path,
    template_file: Path,
    before: Dict[str, int],
    after: Dict[str, int],
    copied_parts: List[str],
    style_stats: Counter,
    table_style_id: Optional[str],
    table_style_stats: Counter,
    cleanup_stats: Counter,
) -> None:
    report.parent.mkdir(parents=True, exist_ok=True)
    rows = [
        ("媒体文件数量", before["media_files"], after["media_files"], "通过" if after["media_files"] >= before["media_files"] else "异常"),
        ("图片关系数量", before["image_relationships"], after["image_relationships"], "通过" if after["image_relationships"] >= before["image_relationships"] else "异常"),
        ("段落数量", before["paragraphs"], after["paragraphs"], "通过" if after["paragraphs"] == before["paragraphs"] else "需复核"),
        ("表格数量", before["tables"], after["tables"], "通过" if after["tables"] == before["tables"] else "异常"),
        ("默认 Normal/正文 样式段落数量", before["normal_style_paragraphs"], after["normal_style_paragraphs"], "通过" if after["normal_style_paragraphs"] <= before["normal_style_paragraphs"] else "需复核"),
        ("自动编号段落数量", before["numbered_paragraphs"], after["numbered_paragraphs"], "通过" if after["numbered_paragraphs"] <= before["numbered_paragraphs"] else "需复核"),
    ]
    with report.open("w", encoding="utf-8") as f:
        f.write("# Word 样式规整校验报告\n\n")
        f.write(f"- 源文件：`{input_file}`\n")
        f.write(f"- 模板文件：`{template_file}`\n")
        f.write(f"- 输出文件：`{output_file}`\n")
        f.write("- 处理方式：基于源 `.docx` 包结构原位处理，不重建正文。\n\n")
        f.write("## 对象与结构校验\n\n")
        f.write("| 检查项 | 原文件 | 新文件 | 结论 |\n")
        f.write("| --- | ---: | ---: | --- |\n")
        for name, b, a, conclusion in rows:
            f.write(f"| {name} | {b} | {a} | {conclusion} |\n")
        f.write("\n## 已复制模板部件\n\n")
        for part in copied_parts:
            f.write(f"- `{part}`\n")
        if not copied_parts:
            f.write("- 未复制模板部件，请检查模板文件。\n")
        f.write("\n## 样式套用统计\n\n")
        f.write("| styleId | 段落数量 |\n")
        f.write("| --- | ---: |\n")
        for sid, count in style_stats.most_common():
            f.write(f"| `{sid}` | {count} |\n")
        f.write("\n## 表格样式套用统计\n\n")
        if table_style_id:
            f.write(f"- 目标表格样式：`表格标准样式` (`{table_style_id}`)\n\n")
            f.write("| table styleId | 表格数量 |\n")
            f.write("| --- | ---: |\n")
            for sid, count in table_style_stats.most_common():
                f.write(f"| `{sid}` | {count} |\n")
            if not table_style_stats:
                f.write(f"| `{table_style_id}` | 0 |\n")
        else:
            f.write("- 未在模板中找到 `表格标准样式`，本次未改写表格本体样式。\n")
        f.write("\n## 直接格式清理统计\n\n")
        f.write("| 清理项 | 数量 |\n")
        f.write("| --- | ---: |\n")
        labels = {
            "paragraph_direct_format": "段落直接格式",
            "run_direct_format": "文字直接格式",
            "table_cell_direct_format": "表格单元格边框/底纹直接格式",
            "table_width_format": "表格总宽/固定布局直接格式",
            "table_column_width_format": "表格列宽直接格式",
        }
        for key, label in labels.items():
            f.write(f"| {label} | {cleanup_stats.get(key, 0)} |\n")
        f.write("\n## 处理结论\n\n")
        if after["media_files"] < before["media_files"] or after["image_relationships"] < before["image_relationships"] or after["tables"] < before["tables"]:
            f.write("本次处理存在对象数量下降，应视为失败输出，请回退源文件重新处理。\n")
        else:
            f.write("本次处理未发现媒体文件、图片关系或表格数量下降。仍建议用 Word 打开输出文件，刷新目录并进行版面抽查。\n")


def default_output_paths(input_file: Path) -> Tuple[Path, Path]:
    return (
        Path("output") / f"{input_file.stem}_标准样式规整.docx",
        Path("output") / f"{input_file.stem}_样式规整校验报告.md",
    )


def normalize_docx(
    input_file: Path,
    template_file: Path,
    output_file: Optional[Path] = None,
    report_file: Optional[Path] = None,
) -> Tuple[Path, Path, Dict[str, int], Dict[str, int]]:
    input_file = Path(input_file)
    template_file = Path(template_file)

    if input_file.suffix.lower() != ".docx":
        raise ValueError("仅支持 .docx 文件；请先将 .doc 转换为 .docx。")
    if template_file.suffix.lower() != ".docx":
        raise ValueError("模板必须为 .docx 文件。")

    default_output, default_report = default_output_paths(input_file)
    output_file = Path(output_file) if output_file else default_output
    report_file = Path(report_file) if report_file else default_report

    before = count_docx(input_file)

    with tempfile.TemporaryDirectory() as td:
        temp = Path(td)
        work = temp / "work"
        template = temp / "template"
        work.mkdir()
        template.mkdir()
        unzip_docx(input_file, work)
        unzip_docx(template_file, template)
        style_hints = parse_style_hints(work / "word/styles.xml")
        numbering = NumberingMaterializer(parse_numbering(work / "word/numbering.xml"))
        table_style_id = parse_table_style_id(template / "word/styles.xml")
        table_look = parse_template_table_look(template / "word/document.xml", table_style_id)
        copied_parts = copy_template_parts(template, work)

        style_stats = Counter()
        table_style_stats = Counter()
        cleanup_stats = Counter()
        for xf in xml_files(work):
            xml_style_stats, xml_table_style_stats, xml_cleanup_stats = normalize_xml(
                xf,
                style_hints,
                numbering,
                table_style_id,
                table_look,
            )
            style_stats.update(xml_style_stats)
            table_style_stats.update(xml_table_style_stats)
            cleanup_stats.update(xml_cleanup_stats)

        zip_dir(work, output_file)

    after = count_docx(output_file)
    generate_report(
        report_file,
        input_file,
        output_file,
        template_file,
        before,
        after,
        copied_parts,
        style_stats,
        table_style_id,
        table_style_stats,
        cleanup_stats,
    )

    if after["media_files"] < before["media_files"] or after["image_relationships"] < before["image_relationships"] or after["tables"] < before["tables"]:
        raise RuntimeError("处理后对象数量下降，已生成报告但不建议交付输出文件。")

    return output_file, report_file, before, after


def main() -> None:
    ap = argparse.ArgumentParser(description="按标准 Word 模板进行原位样式规整")
    ap.add_argument("--input", required=True, type=Path, help="待规整 .docx 文件")
    ap.add_argument("--template", required=True, type=Path, help="标准样式 Word 模板 .docx 文件")
    ap.add_argument("--output", type=Path, help="输出 .docx 文件；默认 output/<源文件名>_标准样式规整.docx")
    ap.add_argument("--report", type=Path, help="校验报告 .md 文件；默认 output/<源文件名>_样式规整校验报告.md")
    args = ap.parse_args()

    if args.input.suffix.lower() != ".docx":
        raise SystemExit("仅支持 .docx 文件；请先将 .doc 转换为 .docx。")
    if args.template.suffix.lower() != ".docx":
        raise SystemExit("模板必须为 .docx 文件。")

    try:
        normalize_docx(args.input, args.template, args.output, args.report)
    except (ValueError, RuntimeError) as exc:
        raise SystemExit(str(exc)) from exc


if __name__ == "__main__":
    main()
