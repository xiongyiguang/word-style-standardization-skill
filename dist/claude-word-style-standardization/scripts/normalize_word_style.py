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
            }
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


def set_pstyle(p: etree._Element, style_id: str) -> None:
    ppr = ensure_ppr(p)
    pstyle = ppr.find(qn("w:pStyle"))
    if pstyle is None:
        pstyle = etree.Element(qn("w:pStyle"))
        ppr.insert(0, pstyle)
    pstyle.set(qn("w:val"), style_id)


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


def bullet_level(p: etree._Element, text: str, numbering: NumberingMaterializer) -> Optional[int]:
    auto_bullet_level = numbering.bullet_level(p)
    if auto_bullet_level:
        return auto_bullet_level

    # 只按文本中的明确项目符号判断。自动编号和缩进只用于编号物化，
    # 不能触发 P2/P3/P4，否则会出现“箭头样式 + 1)”的重复编号。
    if any(pat.match(text) for pat in BULLET_MARKERS):
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

    lvl = bullet_level(p, text, numbering)
    if lvl:
        return STYLE[f"list{lvl}"], True

    if text and (style_hint_contains(style_hints, "bold", current) or is_all_bold(p)) and len(text) <= 80:
        return STYLE["bold_body"], True

    return STYLE["body"], True


def normalize_xml(file: Path, style_hints: Dict[str, object], numbering: NumberingMaterializer) -> Counter:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.parse(str(file), parser).getroot()
    stats = Counter()

    for p in root.xpath(".//w:p", namespaces=NS):
        style_id, remove_numbering = choose_style(p, style_hints, numbering)
        set_pstyle(p, style_id)
        if remove_numbering:
            prefix = numbering.text_for(p)
            if prefix:
                insert_prefix_text(p, prefix)
            remove_numpr(p)
        stats[style_id] += 1

    tree = etree.ElementTree(root)
    tree.write(str(file), xml_declaration=True, encoding="UTF-8", standalone=True)
    return stats


def generate_report(
    report: Path,
    input_file: Path,
    output_file: Path,
    template_file: Path,
    before: Dict[str, int],
    after: Dict[str, int],
    copied_parts: List[str],
    style_stats: Counter,
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
        f.write("\n## 处理结论\n\n")
        if after["media_files"] < before["media_files"] or after["image_relationships"] < before["image_relationships"] or after["tables"] < before["tables"]:
            f.write("本次处理存在对象数量下降，应视为失败输出，请回退源文件重新处理。\n")
        else:
            f.write("本次处理未发现媒体文件、图片关系或表格数量下降。仍建议用 Word 打开输出文件，刷新目录并进行版面抽查。\n")


def main() -> None:
    ap = argparse.ArgumentParser(description="按标准 Word 模板进行原位样式规整")
    ap.add_argument("--input", required=True, type=Path, help="待规整 .docx 文件")
    ap.add_argument("--template", required=True, type=Path, help="标准样式 Word 模板 .docx 文件")
    ap.add_argument("--output", required=True, type=Path, help="输出 .docx 文件")
    ap.add_argument("--report", required=True, type=Path, help="校验报告 .md 文件")
    args = ap.parse_args()

    if args.input.suffix.lower() != ".docx":
        raise SystemExit("仅支持 .docx 文件；请先将 .doc 转换为 .docx。")
    if args.template.suffix.lower() != ".docx":
        raise SystemExit("模板必须为 .docx 文件。")

    before = count_docx(args.input)

    with tempfile.TemporaryDirectory() as td:
        temp = Path(td)
        work = temp / "work"
        template = temp / "template"
        work.mkdir()
        template.mkdir()
        unzip_docx(args.input, work)
        unzip_docx(args.template, template)
        style_hints = parse_style_hints(work / "word/styles.xml")
        numbering = NumberingMaterializer(parse_numbering(work / "word/numbering.xml"))
        copied_parts = copy_template_parts(template, work)

        style_stats = Counter()
        for xf in xml_files(work):
            style_stats.update(normalize_xml(xf, style_hints, numbering))

        zip_dir(work, args.output)

    after = count_docx(args.output)
    generate_report(args.report, args.input, args.output, args.template, before, after, copied_parts, style_stats)

    if after["media_files"] < before["media_files"] or after["image_relationships"] < before["image_relationships"] or after["tables"] < before["tables"]:
        raise SystemExit("处理后对象数量下降，已生成报告但不建议交付输出文件。")


if __name__ == "__main__":
    main()
