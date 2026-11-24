"""Preview and copy page layout, section geometry, numbering, and safe table/indent repairs."""

from __future__ import annotations

import os
from pathlib import Path
import re
from xml.dom import Node, minidom
from xml.etree import ElementTree as ET

from docx_package import W_NS, read_template, resolve_target, same_path, write_package


MAIN = "word/document.xml"
RELS = "word/_rels/document.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
SETTINGS = "word/settings.xml"
SKILL_DIR = Path(__file__).resolve().parent.parent.resolve()
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SETTINGS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
W_ET = f"{{{W_NS}}}"
R_ET = f"{{{R_NS}}}"
SIDES = {"top": "上", "bottom": "下", "left": "左", "right": "右"}
EXTRA_MARGINS = {"header": "页眉距", "footer": "页脚距", "gutter": "装订线"}
ALL_MARGINS = {**SIDES, **EXTRA_MARGINS}
PLACEHOLDER_LABELS = ("此处插入矢量科研图", "在此放置作品、资料或研究示意图", "在此放置高分辨率矢量科研图")
SECT_PR_ORDER = (
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
)
TBL_PR_ORDER = (
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
    "tblDescription", "tblPrChange",
)
TC_PR_ORDER = (
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
    "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
    "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
)


def elem_has_field(root: ET.Element, command: str) -> bool:
    """Detect complex or simple Word fields even when preceded by another field in the same paragraph."""
    pattern = re.compile(r"^\s*" + re.escape(command) + r"(?:\s|$)", re.IGNORECASE)
    if any(pattern.search(field.get(W_ET + "instr", "")) for field in root.iter(W_ET + "fldSimple")):
        return True
    for paragraph in root.iter(W_ET + "p"):
        bufs: list[list[str]] = []
        for node in paragraph.iter():
            if node.tag == W_ET + "fldChar":
                ftype = node.get(W_ET + "fldCharType")
                if ftype == "begin":
                    bufs.append([])
                elif ftype in ("separate", "end") and bufs:
                    field_expr = "".join(bufs.pop())
                    if pattern.search(field_expr) or pattern.search(" ".join(field_expr.split())):
                        return True
            elif node.tag == W_ET + "instrText":
                txt = node.text or ""
                if bufs:
                    bufs[-1].append(txt)
                elif pattern.search(txt):
                    return True
        while bufs:
            field_expr = "".join(bufs.pop())
            if pattern.search(field_expr) or pattern.search(" ".join(field_expr.split())):
                return True
    return False


def _ensure_safe_output(output: Path | None, docx: Path, template: Path) -> None:
    if output is None:
        return
    resolved = output.resolve()
    if resolved == SKILL_DIR or SKILL_DIR in resolved.parents:
        raise PermissionError(f"禁止将修正文件写入技能目录内部：{resolved}")
    if output.suffix.lower() != ".docx" or same_path(output, docx) or same_path(output, template):
        raise ValueError("请为副本指定独立的 .docx 文件名。")
    if output.exists():
        raise FileExistsError(f"副本已存在；请选择新文件名：{output}")


def children(parent, name: str):
    return [
        node
        for node in parent.childNodes
        if node.nodeType == Node.ELEMENT_NODE
        and node.namespaceURI == W_NS
        and node.localName == name
    ]


def _first_child(parent, name: str):
    hits = children(parent, name)
    return hits[0] if hits else None


def _text_of(node) -> str:
    parts: list[str] = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        for child in t.childNodes:
            if child.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE):
                parts.append(child.data)
    return "".join(parts)


def section_margins(document, label: str):
    body = children(document.documentElement, "body")[0]
    sections = document.getElementsByTagNameNS(W_NS, "sectPr")
    final = children(body, "sectPr")
    if len(final) != 1 or not sections or sections[-1] is not final[0]:
        raise ValueError(f"{label} 缺少唯一的末节页面设置，请在 Word 中核对分节。")
    margins = []
    for index, section in enumerate(sections, 1):
        if section is not final[0]:
            props = section.parentNode
            paragraph = props.parentNode
            if not (
                props.namespaceURI == W_NS
                and props.localName == "pPr"
                and paragraph.namespaceURI == W_NS
                and paragraph.localName == "p"
                and paragraph.parentNode is body
            ):
                raise ValueError(f"{label} 第 {index} 处含嵌套或修订分节，请在 Word 中核对分节。")
        candidates = children(section, "pgMar")
        if len(candidates) != 1:
            raise ValueError(f"{label} 第 {index} 节缺少唯一的页边距设置，请在 Word 中核对。")
        margin = candidates[0]
        for side in SIDES:
            raw = margin.getAttributeNS(W_NS, side)
            pattern = r"-?[0-9]+" if side in ("top", "bottom") else r"[0-9]+"
            if not re.fullmatch(pattern, raw):
                raise ValueError(f"{label} 第 {index} 节的{SIDES[side]}边距缺少明确的 Word 整数值。")
        margins.append(margin)
    return margins


def _validated_sections(document, label: str):
    body = children(document.documentElement, "body")[0]
    sections = list(document.getElementsByTagNameNS(W_NS, "sectPr"))
    final = children(body, "sectPr")
    if len(final) != 1 or not sections or sections[-1] is not final[0]:
        raise ValueError(f"{label} 缺少唯一的末节页面设置，请在 Word 中核对分节。")
    for index, section in enumerate(sections, 1):
        if section is not final[0]:
            props = section.parentNode
            paragraph = props.parentNode
            if not (
                props.namespaceURI == W_NS
                and props.localName == "pPr"
                and paragraph.namespaceURI == W_NS
                and paragraph.localName == "p"
                and paragraph.parentNode is body
            ):
                raise ValueError(f"{label} 第 {index} 处含嵌套或修订分节，请在 Word 中核对分节。")
        if section.getElementsByTagNameNS(W_NS, "sectPrChange"):
            raise ValueError(f"{label} 第 {index} 节存在未接受的分节修订，请先在 Word 中接受修订。")
        if len(children(section, "pgMar")) != 1:
            raise ValueError(f"{label} 第 {index} 节缺少唯一的页边距设置，请在 Word 中核对。")
    return sections


def _section_has_page_flags(items: dict[str, bytes]) -> list[bool]:
    """Compute inherited footer PAGE-field presence per section, matching audit_thesis_docx."""
    doc_bytes = items.get(MAIN)
    if not doc_bytes:
        return []
    doc_root = ET.fromstring(doc_bytes)
    rels_map: dict[str, str] = {}
    if RELS in items:
        rels_root = ET.fromstring(items[RELS])
        rels_map = {
            r.get("Id", ""): resolve_target(MAIN, r.get("Target", ""))
            for r in rels_root
            if r.get("TargetMode", "Internal") == "Internal"
        }
    footer_cache: dict[str, bool] = {}
    footers: dict[str, bool] = {}
    flags: list[bool] = []
    for section in doc_root.iter(W_ET + "sectPr"):
        for ref in section.findall(W_ET + "footerReference"):
            role = ref.get(W_ET + "type", "default")
            part = rels_map.get(ref.get(R_ET + "id", ""), "")
            if part not in footer_cache:
                footer_cache[part] = bool(part in items and elem_has_field(ET.fromstring(items[part]), "PAGE"))
            footers[role] = footer_cache[part]
        flags.append(any(footers.values()))
    return flags


def _section_meta_from_dom(section, *, has_page: bool = False) -> dict:
    pg_sz = _first_child(section, "pgSz")
    orient = "portrait"
    w_val = h_val = None
    if pg_sz is not None:
        raw_w = pg_sz.getAttributeNS(W_NS, "w")
        raw_h = pg_sz.getAttributeNS(W_NS, "h")
        w_val = int(raw_w) if re.fullmatch(r"\d+", raw_w) else None
        h_val = int(raw_h) if re.fullmatch(r"\d+", raw_h) else None
        raw_orient = pg_sz.getAttributeNS(W_NS, "orient")
        if raw_orient == "landscape" or (w_val is not None and h_val is not None and w_val > h_val):
            orient = "landscape"
    pg_num = _first_child(section, "pgNumType")
    fmt = (pg_num.getAttributeNS(W_NS, "fmt") or "decimal") if pg_num is not None else "decimal"
    start = (pg_num.getAttributeNS(W_NS, "start") or None) if pg_num is not None else None
    sec_type = _first_child(section, "type")
    kind = (sec_type.getAttributeNS(W_NS, "val") or "nextPage") if sec_type is not None else "nextPage"
    return {
        "orient": orient,
        "w": w_val,
        "h": h_val,
        "fmt": fmt,
        "start": start,
        "kind": kind,
        "has_page": has_page,
    }


def _direct_equal_length_compatible(draft_meta: list[dict], template_meta: list[dict]) -> bool:
    if len(draft_meta) != len(template_meta):
        return False
    for idx, (dm, tm) in enumerate(zip(draft_meta, template_meta)):
        if (dm.get("orient") == "landscape") != (tm.get("orient") == "landscape"):
            return False
        if len(draft_meta) > 1 and idx == 0 and bool(dm.get("has_page")) != bool(tm.get("has_page")):
            return False
    return True


def align_section_indices(draft_meta: list[dict], template_meta: list[dict]) -> list[dict]:
    """Align draft sections onto template sections by semantic role while protecting landscape tables."""
    if not draft_meta or not template_meta:
        raise ValueError("文稿或模板缺少可对齐的分节。")

    if _direct_equal_length_compatible(draft_meta, template_meta):
        direct_pairs = []
        t_dec_idx = next(
            (
                i
                for i, m in enumerate(template_meta)
                if m.get("fmt") == "decimal"
                and m.get("orient") != "landscape"
                and (m.get("start") == "1" or m.get("has_page") or i > 0 or len(template_meta) == 1)
            ),
            len(template_meta) - 1,
        )
        for idx, (dm, tm) in enumerate(zip(draft_meta, template_meta)):
            preserve_landscape = dm.get("orient") == "landscape" and tm.get("orient") != "landscape"
            if preserve_landscape:
                role = "landscape_island"
            elif tm.get("fmt") in ("upperRoman", "lowerRoman"):
                role = "roman_front"
            elif idx < t_dec_idx and not tm.get("has_page") and dm.get("fmt") == "decimal":
                role = "unnumbered_front"
            else:
                role = "decimal_body"
            direct_pairs.append({
                "draft_idx": idx,
                "template_idx": idx,
                "preserve_landscape": preserve_landscape,
                "is_first_body": idx == t_dec_idx,
                "role": role,
            })
        return direct_pairs

    t_landscape = next((i for i, m in enumerate(template_meta) if m.get("orient") == "landscape"), None)
    t_roman = next((i for i, m in enumerate(template_meta) if m.get("fmt") in ("upperRoman", "lowerRoman")), None)
    t_decimal = next(
        (i for i, m in enumerate(template_meta) if m.get("fmt") == "decimal" and m.get("orient") != "landscape" and (i > 0 or len(template_meta) == 1)),
        None,
    )
    if t_decimal is None:
        t_decimal = next((i for i in range(len(template_meta) - 1, -1, -1) if template_meta[i].get("orient") != "landscape"), 0)
    t_front = 0 if (t_roman is not None and t_roman > 0) or (len(template_meta) >= 2 and t_decimal > 0) else t_decimal

    first_roman_or_body = next(
        (
            i
            for i, m in enumerate(draft_meta)
            if m.get("fmt") in ("upperRoman", "lowerRoman")
            or m.get("has_page")
            or m.get("start") == "1"
            or (i > 0 and m.get("orient") != "landscape" and i >= len(draft_meta) - 1)
        ),
        len(draft_meta) - 1,
    )
    if first_roman_or_body == 0 and len(draft_meta) > 1 and t_front != t_decimal:
        first_roman_or_body = 1

    mappings: list[dict] = []
    first_body_assigned = False
    passed_roman = False
    for i, dm in enumerate(draft_meta):
        if dm.get("orient") == "landscape":
            t_idx = t_landscape if t_landscape is not None else t_decimal
            preserve_land = t_landscape is None
            mappings.append({
                "draft_idx": i,
                "template_idx": t_idx,
                "preserve_landscape": preserve_land,
                "is_first_body": False,
                "role": "landscape_island",
            })
            continue
        if dm.get("fmt") in ("upperRoman", "lowerRoman") or (
            t_roman is not None and not passed_roman and i >= first_roman_or_body and i < len(draft_meta) - 1 and not (dm.get("fmt") == "decimal" and dm.get("start") == "1")
        ):
            passed_roman = True
            t_idx = t_roman if t_roman is not None else t_front
            mappings.append({
                "draft_idx": i,
                "template_idx": t_idx,
                "preserve_landscape": False,
                "is_first_body": False,
                "role": "roman_front",
            })
            continue
        if i < first_roman_or_body and not dm.get("has_page"):
            mappings.append({
                "draft_idx": i,
                "template_idx": t_front,
                "preserve_landscape": False,
                "is_first_body": False,
                "role": "unnumbered_front",
            })
            continue
        passed_roman = True
        is_first = not first_body_assigned
        first_body_assigned = True
        mappings.append({
            "draft_idx": i,
            "template_idx": t_decimal,
            "preserve_landscape": False,
            "is_first_body": is_first,
            "role": "decimal_body",
        })
    return mappings


def _fmt_dxa_cm(val: int) -> str:
    return f"{val * 2.54 / 1440:.2f} cm ({val} twips)"


def _w_prefix(doc) -> str:
    root = doc.documentElement if doc is not None else None
    if root is not None and root.attributes is not None:
        for idx in range(root.attributes.length):
            attr = root.attributes.item(idx)
            if attr.value == W_NS and attr.name.startswith("xmlns:"):
                return attr.name.split(":", 1)[1]
    if root is not None:
        # Always bind explicit xmlns:w so setAttributeNS writes prefixed attributes even under default xmlns.
        root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:w", W_NS)
    return "w"


def _w_qname(doc, local: str) -> str:
    prefix = _w_prefix(doc)
    return f"{prefix}:{local}" if prefix else local


def _set_w_attr(el, name: str, value: str) -> None:
    node = el.getAttributeNodeNS(W_NS, name)
    if node is not None:
        node.value = value
    else:
        el.setAttributeNS(W_NS, _w_qname(el.ownerDocument, name), value)


def _remove_w_attr(el, name: str) -> None:
    node = el.getAttributeNodeNS(W_NS, name)
    if node is not None:
        el.removeAttributeNode(node)


def _insert_in_schema_order(parent, new_el, ordered_tags: tuple[str, ...]) -> None:
    try:
        target_pos = ordered_tags.index(new_el.localName)
    except ValueError:
        parent.appendChild(new_el)
        return
    for child in list(parent.childNodes):
        if child.nodeType == Node.ELEMENT_NODE and child.namespaceURI == W_NS:
            try:
                if ordered_tags.index(child.localName) > target_pos:
                    parent.insertBefore(new_el, child)
                    return
            except ValueError:
                continue
    parent.appendChild(new_el)


def _insert_after_margin(section, new_el) -> None:
    _insert_in_schema_order(section, new_el, SECT_PR_ORDER)


def _register_settings_part(draft_items: dict[str, bytes]) -> None:
    if CONTENT_TYPES in draft_items:
        with minidom.parseString(draft_items[CONTENT_TYPES]) as ct_doc:
            ct_root = ct_doc.documentElement
            has_ov = any(
                node.getAttribute("PartName") == "/word/settings.xml"
                for node in ct_root.getElementsByTagNameNS(CT_NS, "Override")
            )
            if not has_ov:
                prefix = ct_root.prefix
                ov = ct_doc.createElementNS(CT_NS, f"{prefix}:Override" if prefix else "Override")
                ov.setAttribute("PartName", "/word/settings.xml")
                ov.setAttribute("ContentType", SETTINGS_CT)
                ct_root.appendChild(ov)
                draft_items[CONTENT_TYPES] = ct_doc.toxml(encoding="UTF-8")
    if RELS in draft_items:
        with minidom.parseString(draft_items[RELS]) as rel_doc:
            rel_root = rel_doc.documentElement
            rels = rel_root.getElementsByTagNameNS(REL_NS, "Relationship")
            has_rel = any(r.getAttribute("Type") == SETTINGS_REL_TYPE for r in rels)
            if not has_rel:
                used = {r.getAttribute("Id") for r in rels}
                n = 1
                while f"rId{n}" in used:
                    n += 1
                prefix = rel_root.prefix
                rel = rel_doc.createElementNS(REL_NS, f"{prefix}:Relationship" if prefix else "Relationship")
                rel.setAttribute("Id", f"rId{n}")
                rel.setAttribute("Type", SETTINGS_REL_TYPE)
                rel.setAttribute("Target", "settings.xml")
                rel_root.appendChild(rel)
                draft_items[RELS] = rel_doc.toxml(encoding="UTF-8")


def _sync_settings(draft_items: dict[str, bytes], template_items: dict[str, bytes]) -> list[str]:
    want: dict[str, bool] = {"mirrorMargins": False, "evenAndOddHeaders": False}
    if SETTINGS in template_items:
        with minidom.parseString(template_items[SETTINGS]) as t_doc:
            t_root = t_doc.documentElement
            for flag in want:
                t_node = _first_child(t_root, flag)
                want[flag] = t_node is not None and (t_node.getAttributeNS(W_NS, "val") or "1") not in ("0", "false", "off")
    if SETTINGS not in draft_items and not any(want.values()):
        return []
    changes: list[str] = []
    if SETTINGS not in draft_items:
        draft_items[SETTINGS] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        _register_settings_part(draft_items)
        changes.append("全局设置 · 补建 word/settings.xml 部件")
    with minidom.parseString(draft_items[SETTINGS]) as d_doc:
        d_root = d_doc.documentElement
        dirty = False
        for flag, label in (("mirrorMargins", "对称页边距"), ("evenAndOddHeaders", "奇偶页不同页眉")):
            t_on = want[flag]
            d_node = _first_child(d_root, flag)
            d_on = d_node is not None and (d_node.getAttributeNS(W_NS, "val") or "1") not in ("0", "false", "off")
            if t_on == d_on:
                continue
            if t_on:
                if d_node is None:
                    el = d_doc.createElementNS(W_NS, _w_qname(d_doc, flag))
                    if d_root.firstChild is not None:
                        d_root.insertBefore(el, d_root.firstChild)
                    else:
                        d_root.appendChild(el)
                else:
                    _remove_w_attr(d_node, "val")
                dirty = True
                changes.append(f"全局设置 · 开启{label} ({flag})")
            else:
                if d_node is not None:
                    d_root.removeChild(d_node)
                dirty = True
                changes.append(f"全局设置 · 关闭{label} ({flag})")
        if dirty:
            draft_items[SETTINGS] = d_doc.toxml(encoding="UTF-8")
    return changes


def _is_placeholder_table(tbl) -> bool:
    bdr = None
    tbl_pr = _first_child(tbl, "tblPr")
    if tbl_pr is not None:
        bdr = _first_child(tbl_pr, "tblBorders")
    top_b = _first_child(bdr, "top") if bdr is not None else None
    rows = children(tbl, "tr")
    cells = [c for r in rows for c in children(r, "tc")]
    return (
        len(cells) == 1
        and top_b is not None
        and top_b.getAttributeNS(W_NS, "val") == "dashed"
        and _text_of(tbl).lstrip("【[ ").startswith(PLACEHOLDER_LABELS)
    )


def _is_front_heading_text(txt: str) -> bool:
    return bool(
        re.match(
            r"^(?:摘\s*要|中文摘要|abstract|目\s*录|contents|版权.*声明|致\s*谢|个人简历|附\s*录|参考\s*文献)\b",
            txt.strip(),
            re.I,
        )
    )


def _is_toc_dom_para(p) -> bool:
    ppr = _first_child(p, "pPr")
    if ppr is not None:
        ps_el = _first_child(ppr, "pStyle")
        if ps_el is not None and ps_el.getAttributeNS(W_NS, "val").lower().startswith("toc"):
            return True
    for el in p.getElementsByTagNameNS(W_NS, "fldSimple"):
        if re.match(r"^\s*(?:TOC|PAGEREF)\b", el.getAttributeNS(W_NS, "instr"), re.I):
            return True
    for el in p.getElementsByTagNameNS(W_NS, "instrText"):
        txt = "".join(c.data for c in el.childNodes if c.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE))
        if re.match(r"^\s*(?:TOC|PAGEREF)\b", txt, re.I):
            return True
    return False


def _is_chapter1_dom_para(p) -> bool:
    if _is_toc_dom_para(p):
        return False
    txt = _text_of(p).strip()
    if not txt or _is_front_heading_text(txt):
        return False
    ppr = _first_child(p, "pPr")
    ps_el = _first_child(ppr, "pStyle") if ppr is not None else None
    ps_val = ps_el.getAttributeNS(W_NS, "val") if ps_el is not None else ""
    if re.match(r"^(?:第一章|Chapter\s+1)\b", txt, re.I):
        return True
    if ps_val in ("Heading1Chapter", "04"):
        return True
    if ps_val in ("Heading1", "heading 1") and ppr is not None and _first_child(ppr, "numPr") is not None:
        return True
    return False


def _iter_dom_story_blocks(container):
    for child in list(container.childNodes):
        if child.nodeType != Node.ELEMENT_NODE or child.namespaceURI != W_NS:
            continue
        if child.localName in ("del", "moveFrom", "txbxContent"):
            continue
        if child.localName == "sdt":
            sdt_c = _first_child(child, "sdtContent")
            if sdt_c is not None:
                yield from _iter_dom_story_blocks(sdt_c)
        elif child.localName in ("p", "tbl"):
            yield child


def _normalize_three_line_tables(document) -> list[str]:
    changes: list[str] = []
    body = children(document.documentElement, "body")[0]
    blocks = list(_iter_dom_story_blocks(body))
    total_sections = len(document.getElementsByTagNameNS(W_NS, "sectPr"))
    has_ch1_anywhere = any(b.localName == "p" and _is_chapter1_dom_para(b) for b in blocks)
    seen_ch1 = False
    sec_idx = 0
    t_idx = 0
    for child in blocks:
        if child.localName == "p":
            if _is_chapter1_dom_para(child):
                seen_ch1 = True
            ppr = _first_child(child, "pPr")
            if ppr is not None and _first_child(ppr, "sectPr") is not None:
                sec_idx += 1
        elif child.localName == "tbl":
            if _is_placeholder_table(child):
                continue
            if total_sections > 1 and sec_idx == 0:
                continue
            if has_ch1_anywhere and not seen_ch1:
                continue
            if not seen_ch1 and any(k in _text_of(child) for k in ("姓名", "学号", "导师", "题目", "专业")):
                continue
            t_idx += 1
            rows = children(child, "tr")
            if not rows:
                continue
            tbl_pr = _first_child(child, "tblPr")
            if tbl_pr is None:
                tbl_pr = document.createElementNS(W_NS, _w_qname(document, "tblPr"))
                child.insertBefore(tbl_pr, child.firstChild)
            bdr = _first_child(tbl_pr, "tblBorders")
            if bdr is None:
                bdr = document.createElementNS(W_NS, _w_qname(document, "tblBorders"))
                _insert_in_schema_order(tbl_pr, bdr, TBL_PR_ORDER)
            for existing in list(bdr.childNodes):
                if existing.nodeType == Node.ELEMENT_NODE:
                    bdr.removeChild(existing)
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                el = document.createElementNS(W_NS, _w_qname(document, side))
                if side in ("top", "bottom"):
                    _set_w_attr(el, "val", "single")
                    _set_w_attr(el, "sz", "12")
                    _set_w_attr(el, "space", "0")
                    _set_w_attr(el, "color", "auto")
                else:
                    _set_w_attr(el, "val", "none")
                bdr.appendChild(el)

            # Support single- and multi-row headers (<w:tblHeader/> on leading rows)
            header_rows = 0
            for row in rows:
                tr_pr_check = _first_child(row, "trPr")
                th_node = _first_child(tr_pr_check, "tblHeader") if tr_pr_check is not None else None
                if th_node is not None and (th_node.getAttributeNS(W_NS, "val") or "1") not in ("0", "false", "off"):
                    header_rows += 1
                else:
                    break
            if header_rows == 0:
                header_rows = 1
            for h_i in range(header_rows):
                h_row = rows[h_i]
                tr_pr = _first_child(h_row, "trPr")
                if tr_pr is None:
                    tr_pr = document.createElementNS(W_NS, _w_qname(document, "trPr"))
                    h_row.insertBefore(tr_pr, h_row.firstChild)
                if _first_child(tr_pr, "tblHeader") is None:
                    tr_pr.appendChild(document.createElementNS(W_NS, _w_qname(document, "tblHeader")))
                if _first_child(tr_pr, "cantSplit") is None:
                    tr_pr.appendChild(document.createElementNS(W_NS, _w_qname(document, "cantSplit")))

            for r_i, row in enumerate(rows):
                for cell in children(row, "tc"):
                    tc_pr = _first_child(cell, "tcPr")
                    if tc_pr is None:
                        tc_pr = document.createElementNS(W_NS, _w_qname(document, "tcPr"))
                        cell.insertBefore(tc_pr, cell.firstChild)
                    cb = _first_child(tc_pr, "tcBorders")
                    if r_i == header_rows - 1:
                        if cb is None:
                            cb = document.createElementNS(W_NS, _w_qname(document, "tcBorders"))
                            _insert_in_schema_order(tc_pr, cb, TC_PR_ORDER)
                        for existing in list(cb.childNodes):
                            if existing.nodeType == Node.ELEMENT_NODE:
                                cb.removeChild(existing)
                        bot = document.createElementNS(W_NS, _w_qname(document, "bottom"))
                        _set_w_attr(bot, "val", "single")
                        _set_w_attr(bot, "sz", "6")
                        _set_w_attr(bot, "space", "0")
                        _set_w_attr(bot, "color", "auto")
                        cb.appendChild(bot)
                    elif r_i >= header_rows and cb is not None:
                        tc_pr.removeChild(cb)
            changes.append(f"三线表 #{t_idx} · 已规范 1.5 pt 上下外框、0.75 pt 表头线并清除竖线")
    return changes


def _clean_paragraph_leading_spaces(document) -> list[str]:
    changes: list[str] = []
    body = children(document.documentElement, "body")[0]
    cleaned_count = 0
    non_content = {"pPr", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "proofErr", "permStart", "permEnd"}
    non_text_run_items = ("fldChar", "drawing", "pict", "object", "sym", "br")
    for p in body.getElementsByTagNameNS(W_NS, "p"):
        ppr = _first_child(p, "pPr")
        if ppr is None:
            continue
        ind = _first_child(ppr, "ind")
        explicit_zero_override = False
        positive_indent = False
        if ind is not None:
            fl = ind.getAttributeNS(W_NS, "firstLine")
            flc = ind.getAttributeNS(W_NS, "firstLineChars")
            for raw_val in (fl, flc):
                if raw_val:
                    if re.fullmatch(r"-?\d+", raw_val):
                        if int(raw_val) > 0:
                            positive_indent = True
                        else:
                            explicit_zero_override = True
            if positive_indent:
                explicit_zero_override = False
        if explicit_zero_override:
            continue
        ps_el = _first_child(ppr, "pStyle")
        ps = ps_el.getAttributeNS(W_NS, "val") if ps_el is not None else ""
        if not positive_indent and ps not in ("BodyTextZH", "08", "AbstractBodyZH", "02"):
            continue
        for child in list(p.childNodes):
            if child.nodeType != Node.ELEMENT_NODE:
                continue
            if child.namespaceURI != W_NS:
                break
            if child.localName in non_content:
                continue
            if child.localName != "r":
                break
            if any(_first_child(child, tag) is not None for tag in non_text_run_items):
                break
            tabs = children(child, "tab")
            for tab in tabs:
                child.removeChild(tab)
                cleaned_count += 1
            remaining_text = False
            for t_node in children(child, "t"):
                for tn in list(t_node.childNodes):
                    if tn.nodeType == Node.TEXT_NODE and tn.data:
                        stripped = tn.data.lstrip(" \u3000\t")
                        if stripped != tn.data:
                            tn.data = stripped
                            cleaned_count += 1
                        if tn.data:
                            remaining_text = True
                        break
                if remaining_text:
                    break
            if remaining_text:
                break
    if cleaned_count:
        changes.append(f"段落缩进 · 已清理 {cleaned_count} 处与首行缩进重复的手敲前导空格/制表符")
    return changes


def repair(docx: Path, template: Path, output: Path | None) -> None:
    if docx.suffix.lower() != ".docx":
        raise ValueError("请选择现有论文 .docx 作为文稿。")
    _ensure_safe_output(output, docx, template)

    draft, _ = read_template(docx)
    model, _ = read_template(template)
    with minidom.parseString(draft[MAIN]) as document, minidom.parseString(model[MAIN]) as school:
        actual = section_margins(document, "文稿")
        expected = section_margins(school, "模板")
        if len(actual) != len(expected):
            raise ValueError(
                f"文稿有 {len(actual)} 节，模板有 {len(expected)} 节。请在 Word 中核对两者的分节对应关系，或改用支持跨节映射的 `repair` 命令。"
            )
        changes = []
        for index, (target, reference) in enumerate(zip(actual, expected), 1):
            for side, label in SIDES.items():
                attr = target.getAttributeNodeNS(W_NS, side)
                old = int(attr.value)
                new = int(reference.getAttributeNS(W_NS, side))
                if old != new:
                    attr.value = str(new)
                    changes.append(
                        f"第 {index} 节 · {label}边距 ({side}): "
                        f"{old * 2.54 / 1440:.2f} → {new * 2.54 / 1440:.2f} cm "
                        f"({old} → {new} twips)"
                    )
        if output is not None:
            if changes:
                draft[MAIN] = document.toxml(encoding="UTF-8")
            write_package(output, draft, overwrite=False)
    for line in changes:
        print(line)
    print(f"SECTIONS={len(actual)} MARGIN_CHANGES={len(changes)}")
    if output is None:
        print("预览完成。生成修正版请在同一命令后加 --output 新文件名.docx。")
    else:
        print(f"SAVED={output.resolve()}")


def repair_layout(
    docx: Path,
    template: Path,
    output: Path | None,
    *,
    fix_numbering: bool = True,
    fix_tables: bool = False,
    clean_leading_spaces: bool = False,
) -> None:
    """Preview or repair full section geometry, page-numbering restarts, settings, tables, and stray leading spaces."""
    if docx.suffix.lower() != ".docx":
        raise ValueError("请选择现有论文 .docx 作为文稿。")
    _ensure_safe_output(output, docx, template)

    draft, _ = read_template(docx)
    model, _ = read_template(template)
    d_page_flags = _section_has_page_flags(draft)
    t_page_flags = _section_has_page_flags(model)
    changes: list[str] = []
    main_dirty = False

    with minidom.parseString(draft[MAIN]) as document, minidom.parseString(model[MAIN]) as school:
        d_sections = _validated_sections(document, "文稿")
        t_sections = _validated_sections(school, "模板")
        d_meta = [
            _section_meta_from_dom(s, has_page=d_page_flags[i] if i < len(d_page_flags) else False)
            for i, s in enumerate(d_sections)
        ]
        t_meta = [
            _section_meta_from_dom(s, has_page=t_page_flags[i] if i < len(t_page_flags) else False)
            for i, s in enumerate(t_sections)
        ]
        mappings = align_section_indices(d_meta, t_meta)

        for item in mappings:
            di = item["draft_idx"]
            ti = item["template_idx"]
            d_sec = d_sections[di]
            t_sec = t_sections[ti]
            d_mar = _first_child(d_sec, "pgMar")
            t_mar = _first_child(t_sec, "pgMar")

            # Sync page size & orientation (protecting landscape wide-table islands)
            d_sz = _first_child(d_sec, "pgSz")
            t_sz = _first_child(t_sec, "pgSz")
            if t_sz is not None:
                tw_raw = t_sz.getAttributeNS(W_NS, "w")
                th_raw = t_sz.getAttributeNS(W_NS, "h")
                if re.fullmatch(r"\d+", tw_raw) and re.fullmatch(r"\d+", th_raw):
                    tw, th = int(tw_raw), int(th_raw)
                    target_w, target_h = (max(tw, th), min(tw, th)) if item["preserve_landscape"] else (tw, th)
                    if d_sz is None:
                        d_sz = document.createElementNS(W_NS, _w_qname(document, "pgSz"))
                        _insert_in_schema_order(d_sec, d_sz, SECT_PR_ORDER)
                        cur_w = "未设"
                        cur_h = "未设"
                    else:
                        cur_w = d_sz.getAttributeNS(W_NS, "w")
                        cur_h = d_sz.getAttributeNS(W_NS, "h")
                    if cur_w != str(target_w) or cur_h != str(target_h):
                        _set_w_attr(d_sz, "w", str(target_w))
                        _set_w_attr(d_sz, "h", str(target_h))
                        main_dirty = True
                        changes.append(
                            f"第 {di + 1} 节（对应模板第 {ti + 1} 节）· 纸张尺寸: "
                            f"{cur_w}×{cur_h} → {target_w}×{target_h} twips"
                        )
                    target_orient = "landscape" if (item["preserve_landscape"] or target_w > target_h) else "portrait"
                    cur_orient = d_sz.getAttributeNS(W_NS, "orient") or "portrait"
                    if cur_orient != target_orient:
                        if target_orient == "landscape":
                            _set_w_attr(d_sz, "orient", "landscape")
                        else:
                            _remove_w_attr(d_sz, "orient")
                        main_dirty = True
                        changes.append(f"第 {di + 1} 节 · 页面方向: {cur_orient} → {target_orient}")

            # Sync all 7 pgMar attributes (top/bottom/left/right/header/footer/gutter)
            for attr_name, label in ALL_MARGINS.items():
                t_raw = t_mar.getAttributeNS(W_NS, attr_name)
                if not re.fullmatch(r"-?[0-9]+", t_raw):
                    continue
                new_val = int(t_raw)
                d_raw = d_mar.getAttributeNS(W_NS, attr_name)
                old_val = int(d_raw) if re.fullmatch(r"-?[0-9]+", d_raw) else None
                if attr_name == "gutter" and old_val is None and new_val == 0:
                    continue
                if old_val != new_val:
                    _set_w_attr(d_mar, attr_name, str(new_val))
                    main_dirty = True
                    old_str = _fmt_dxa_cm(old_val) if old_val is not None else "未设"
                    changes.append(
                        f"第 {di + 1} 节（对应模板第 {ti + 1} 节）· {label} ({attr_name}): "
                        f"{old_str} → {_fmt_dxa_cm(new_val)}"
                    )

            # Sync pgNumType when enabled and not a split-continuation body section
            if fix_numbering:
                d_num = _first_child(d_sec, "pgNumType")
                t_num = _first_child(t_sec, "pgNumType")
                if t_num is None and item["role"] == "unnumbered_front":
                    if d_num is not None:
                        d_sec.removeChild(d_num)
                        main_dirty = True
                        changes.append(f"第 {di + 1} 节 · 移除封面/前部节多余页码起始设置")
                elif t_num is not None:
                    t_fmt = t_num.getAttributeNS(W_NS, "fmt") or "decimal"
                    t_start = t_num.getAttributeNS(W_NS, "start") or None
                    want_start = (
                        t_start
                        if (item["role"] == "roman_front" or (item["role"] == "decimal_body" and item["is_first_body"]))
                        else None
                    )
                    if d_num is None:
                        d_num = document.createElementNS(W_NS, _w_qname(document, "pgNumType"))
                        _set_w_attr(d_num, "fmt", t_fmt)
                        if want_start is not None:
                            _set_w_attr(d_num, "start", want_start)
                        _insert_after_margin(d_sec, d_num)
                        main_dirty = True
                        changes.append(
                            f"第 {di + 1} 节 · 补全页码格式 fmt={t_fmt}"
                            + (f", start={want_start}" if want_start else "")
                        )
                    else:
                        d_fmt = d_num.getAttributeNS(W_NS, "fmt") or "decimal"
                        d_start = d_num.getAttributeNS(W_NS, "start") or None
                        if d_fmt != t_fmt:
                            _set_w_attr(d_num, "fmt", t_fmt)
                            main_dirty = True
                            changes.append(f"第 {di + 1} 节 · 页码格式: {d_fmt} → {t_fmt}")
                        if d_start != want_start:
                            if want_start is not None:
                                _set_w_attr(d_num, "start", want_start)
                            else:
                                _remove_w_attr(d_num, "start")
                            main_dirty = True
                            changes.append(
                                f"第 {di + 1} 节 · 起始页码: {d_start or '续编'} → {want_start or '续编'}"
                            )

        if fix_tables:
            tbl_changes = _normalize_three_line_tables(document)
            if tbl_changes:
                main_dirty = True
                changes.extend(tbl_changes)

        if clean_leading_spaces:
            sp_changes = _clean_paragraph_leading_spaces(document)
            if sp_changes:
                main_dirty = True
                changes.extend(sp_changes)

        if main_dirty and output is not None:
            draft[MAIN] = document.toxml(encoding="UTF-8")

    settings_changes = _sync_settings(draft, model)
    changes.extend(settings_changes)

    if output is not None:
        write_package(output, draft, overwrite=False)

    for line in changes:
        print(line)
    print(f"SECTIONS_DRAFT={len(d_meta)} SECTIONS_TEMPLATE={len(t_meta)} LAYOUT_CHANGES={len(changes)}")
    if output is None:
        print("预览完成。生成修正版请在同一命令后加 --output 新文件名.docx。")
    else:
        print(f"SAVED={output.resolve()}")
