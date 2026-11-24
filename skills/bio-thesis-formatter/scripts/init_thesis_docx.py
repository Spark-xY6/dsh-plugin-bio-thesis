#!/usr/bin/env python3
"""Initialize a clean, institution-neutral thesis or milestone report DOCX using Python stdlib only."""

from __future__ import annotations

import argparse
import copy
import json
import pathlib
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

from docx_package import fill_template, identity_parts, read_template, same_path, write_package

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
XML_NS = "http://www.w3.org/XML/1998/namespace"

for _pfx, _uri in (
    ("w", W_NS),
    ("r", R_NS),
    ("w14", W14_NS),
    ("wp", WP_NS),
    ("a", A_NS),
    ("pic", PIC_NS),
    ("cp", CP_NS),
    ("dc", DC_NS),
    ("dcterms", DCTERMS_NS),
    ("xsi", XSI_NS),
):
    ET.register_namespace(_pfx, _uri)

W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
W14 = f"{{{W14_NS}}}"

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent
MASTER_DOCX = SKILL_DIR / "resources" / "bio-thesis-master.docx"
PRESETS_DIR = SKILL_DIR / "resources" / "presets"

_PARA_SEQ = 0x20000000


def _next_id() -> str:
    global _PARA_SEQ
    _PARA_SEQ += 1
    return f"{_PARA_SEQ:08X}"


def _p_text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(f"{W}t"))


def _deep_merge(base: dict, override: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def load_effective_preset(
    preset_name: str,
    profile_config: str | None,
    university: str | None,
    school: str | None,
    citation_style: str,
) -> dict:
    preset_path = PRESETS_DIR / f"{preset_name}.json"
    if not preset_path.exists():
        raise FileNotFoundError(f"Preset not found: {preset_path}")
    cfg = json.loads(preset_path.read_text(encoding="utf-8"))
    if profile_config:
        override = json.loads(pathlib.Path(profile_config).read_text(encoding="utf-8"))
        cfg = _deep_merge(cfg, override)
        # An explicit dxa override must not be shadowed by an inherited cm value.
        page_override = override.get('page', {})
        for key in page_override.get('margins_dxa', {}):
            if key not in page_override.get('margins_cm', {}):
                cfg['page'].get('margins_cm', {}).pop(key, None)
    if university:
        cfg["university_name"] = university
    if school:
        cfg["school_name"] = school
    if citation_style != "auto":
        cfg["citation_style"] = citation_style
    return cfg


def _make_run(
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    superscript: bool = False,
    ascii_font: str | None = None,
    ea_font: str | None = None,
    sz: int | None = None,
) -> ET.Element:
    r = ET.Element(f"{W}r")
    if bold or italic or superscript or ascii_font or ea_font or sz is not None:
        rpr = ET.SubElement(r, f"{W}rPr")
        if ascii_font or ea_font:
            fa: dict[str, str] = {}
            if ascii_font:
                fa[f"{W}ascii"] = ascii_font
                fa[f"{W}hAnsi"] = ascii_font
                fa[f"{W}cs"] = ascii_font
            if ea_font:
                fa[f"{W}eastAsia"] = ea_font
            ET.SubElement(rpr, f"{W}rFonts", fa)
        if bold:
            ET.SubElement(rpr, f"{W}b")
            ET.SubElement(rpr, f"{W}bCs")
        if italic:
            ET.SubElement(rpr, f"{W}i")
            ET.SubElement(rpr, f"{W}iCs")
        if sz is not None:
            ET.SubElement(rpr, f"{W}sz", {f"{W}val": str(sz)})
            ET.SubElement(rpr, f"{W}szCs", {f"{W}val": str(sz)})
        if superscript:
            ET.SubElement(rpr, f"{W}vertAlign", {f"{W}val": "superscript"})
    t = ET.SubElement(r, f"{W}t")
    t.text = text
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set(f"{{{XML_NS}}}space", "preserve")
    return r


def _make_p(text_or_runs: str | list[ET.Element], style: str, *, bold: bool = False) -> ET.Element:
    p = ET.Element(f"{W}p", {f"{W14}paraId": _next_id(), f"{W14}textId": "77777777"})
    ppr = ET.SubElement(p, f"{W}pPr")
    ET.SubElement(ppr, f"{W}pStyle", {f"{W}val": style})
    if isinstance(text_or_runs, str):
        p.append(_make_run(text_or_runs, bold=bold))
    else:
        for r in text_or_runs:
            p.append(r)
    return p


def _cm_to_dxa(cm: float) -> int:
    return int(round(cm * 1440.0 / 2.54))


def _resolve_page_dxa(page_cfg: dict) -> dict[str, int]:
    dxa = dict(page_cfg.get("margins_dxa") or {})
    cm = page_cfg.get("margins_cm") or {}
    defaults = {"top": 1701, "bottom": 1418, "left": 1474, "right": 1474, "header": 1134, "footer": 992, "gutter": 0}
    out: dict[str, int] = {}
    for k, dv in defaults.items():
        if k in cm:
            out[k] = _cm_to_dxa(float(cm[k]))
        else:
            out[k] = int(dxa.get(k, dv))
    out["width_dxa"] = int(page_cfg.get("width_dxa", 11906))
    out["height_dxa"] = int(page_cfg.get("height_dxa", 16838))
    return out


def _read_image_dims(img_path: pathlib.Path) -> tuple[str, str, int, int]:
    data = img_path.read_bytes()
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        w, h = struct.unpack(">II", data[16:24])
        return "png", "image/png", w, h
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                break
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2):
                h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                return "jpg", "image/jpeg", w, h
            if marker in (0xD8, 0xD9):
                i += 2
                continue
            seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
            i += 2 + seg_len
        return "jpg", "image/jpeg", 400, 400
    raise ValueError(f"Unsupported --logo format (only PNG or JPEG supported): {img_path}")


def _make_logo_paragraph(rel_id: str, ext: str, px_w: int, px_h: int) -> ET.Element:
    target_h = 792000  # ~2.2 cm
    target_w = max(360000, min(2800000, int(round(target_h * (px_w / max(px_h, 1))))))
    p = ET.Element(f"{W}p", {f"{W14}paraId": _next_id(), f"{W14}textId": "77777777"})
    ppr = ET.SubElement(p, f"{W}pPr")
    sp = ET.SubElement(ppr, f"{W}spacing", {f"{W}before": "120", f"{W}after": "120"})
    _ = sp
    ET.SubElement(ppr, f"{W}jc", {f"{W}val": "center"})
    r = ET.SubElement(p, f"{W}r")
    drawing = ET.SubElement(r, f"{W}drawing")
    inline = ET.SubElement(drawing, f"{{{WP_NS}}}inline", {"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
    ET.SubElement(inline, f"{{{WP_NS}}}extent", {"cx": str(target_w), "cy": str(target_h)})
    ET.SubElement(inline, f"{{{WP_NS}}}effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
    ET.SubElement(inline, f"{{{WP_NS}}}docPr", {"id": "1", "name": f"UniversityLogo.{ext}"})
    c_nv = ET.SubElement(inline, f"{{{WP_NS}}}cNvGraphicFramePr")
    ET.SubElement(c_nv, f"{{{A_NS}}}graphicFrameLocks", {"noChangeAspect": "1"})
    graphic = ET.SubElement(inline, f"{{{A_NS}}}graphic")
    gdata = ET.SubElement(graphic, f"{{{A_NS}}}graphicData", {"uri": PIC_NS})
    pic = ET.SubElement(gdata, f"{{{PIC_NS}}}pic")
    nv_pic = ET.SubElement(pic, f"{{{PIC_NS}}}nvPicPr")
    ET.SubElement(nv_pic, f"{{{PIC_NS}}}cNvPr", {"id": "1", "name": f"logo.{ext}"})
    ET.SubElement(nv_pic, f"{{{PIC_NS}}}cNvPicPr")
    blip_fill = ET.SubElement(pic, f"{{{PIC_NS}}}blipFill")
    ET.SubElement(blip_fill, f"{{{A_NS}}}blip", {f"{R}embed": rel_id})
    stretch = ET.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    ET.SubElement(stretch, f"{{{A_NS}}}fillRect")
    sp_pr = ET.SubElement(pic, f"{{{PIC_NS}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{A_NS}}}xfrm")
    ET.SubElement(xfrm, f"{{{A_NS}}}off", {"x": "0", "y": "0"})
    ET.SubElement(xfrm, f"{{{A_NS}}}ext", {"cx": str(target_w), "cy": str(target_h)})
    prst = ET.SubElement(sp_pr, f"{{{A_NS}}}prstGeom", {"prst": "rect"})
    ET.SubElement(prst, f"{{{A_NS}}}avLst")
    return p


def _patch_cover(
    body: ET.Element,
    preset: dict,
    mode_cfg: dict,
    *,
    title_zh: str | None,
    author: str | None,
    student_id: str | None,
    major: str | None,
    advisor: str | None,
    date_zh: str | None,
    blind: bool,
    logo_info: tuple[str, str, int, int] | None,
) -> None:
    univ = preset.get("university_name", "【学校名称】")
    school = preset.get("school_name", "【培养单位】")
    cover_lines = [line.format(university=univ, school=school) for line in mode_cfg["cover_lines"]]

    children = list(body)
    # p0 = university line, p1 = degree/report title line
    if children:
        t0 = children[0].find(f".//{W}t")
        if t0 is not None:
            t0.text = univ
    if len(children) >= 2:
        p1 = children[1]
        tmpl = copy.deepcopy(p1)
        pos = list(body).index(p1)
        body.remove(p1)
        for off, line in enumerate(cover_lines):
            new_p = copy.deepcopy(tmpl)
            new_p.set(f"{W14}paraId", _next_id())
            t_el = new_p.find(f".//{W}t")
            if t_el is not None:
                t_el.text = line
            rpr = new_p.find(f".//{W}rPr")
            if len(cover_lines) > 1 and rpr is not None:
                sz = rpr.find(f"{W}sz")
                sz_cs = rpr.find(f"{W}szCs")
                if sz is not None:
                    sz.set(f"{W}val", "48")
                if sz_cs is not None:
                    sz_cs.set(f"{W}val", "48")
            body.insert(pos + off, new_p)

    if logo_info is not None:
        rel_id, ext, px_w, px_h = logo_info
        body.insert(0, _make_logo_paragraph(rel_id, ext, px_w, px_h))

    tbls = list(body.findall(f"{W}tbl"))
    if title_zh and tbls:
        for t_el in tbls[0].iter(f"{W}t"):
            if "【论文中文题目】" in (t_el.text or ""):
                t_el.text = title_zh

    if len(tbls) >= 2:
        meta_tbl = tbls[1]
        for tr in list(meta_tbl.findall(f"{W}tr")):
            cells = tr.findall(f"{W}tc")
            if len(cells) != 2:
                continue
            key_text = _p_text(cells[0]).replace(" ", "").replace("：", "")
            if not mode_cfg["full_cover_fields"] and key_text in ("培养单位", "院系", "研究方向"):
                meta_tbl.remove(tr)
                continue
            val_t = cells[1].find(f".//{W}t")
            if val_t is None:
                continue
            if blind and key_text in ("姓名", "学号", "导师姓名"):
                val_t.text = "【匿名评审隐去】"
            elif key_text in ("培养单位", "院系"):
                val_t.text = school
            elif key_text == "姓名" and author:
                val_t.text = author
            elif key_text == "学号" and student_id:
                val_t.text = student_id
            elif key_text == "专业" and major:
                val_t.text = major
            elif key_text == "导师姓名" and advisor:
                val_t.text = advisor

    if date_zh:
        for p in body.findall(f"{W}p"):
            if "二〇二X" in _p_text(p):
                t_el = p.find(f".//{W}t")
                if t_el is not None:
                    t_el.text = date_zh
                break


def _patch_front_and_back(
    body: ET.Element,
    mode_cfg: dict,
    *,
    title_en: str | None,
    author: str | None,
    major: str | None,
    advisor: str | None,
    univ: str | None,
    school: str | None,
    blind: bool,
) -> None:
    children = list(body)
    if not mode_cfg["include_copyright"]:
        to_remove = []
        for idx, el in enumerate(children):
            txt = _p_text(el).strip()
            if txt == "版权与使用声明":
                if idx > 0 and children[idx - 1].find(f".//{W}br[@{W}type='page']") is not None:
                    to_remove.append(children[idx - 1])
                to_remove.append(el)
                if idx + 1 < len(children) and "【通用版权声明" in _p_text(children[idx + 1]):
                    to_remove.append(children[idx + 1])
                break
        for el in to_remove:
            body.remove(el)

    if not mode_cfg.get("include_cv", True):
        children = list(body)
        cv_idx = next((i for i, el in enumerate(children) if _p_text(el).strip().startswith(("个人简历", "作者简历"))), None)
        if cv_idx is not None:
            cv_end = next(
                (
                    i
                    for i in range(cv_idx + 1, len(children))
                    if children[i].tag == f"{W}p"
                    and _p_text(children[i]).strip().replace("\u3000", "").replace(" ", "") in ("致谢", "学位论文独创性与使用授权声明")
                ),
                len(children),
            )
            for el in children[cv_idx:cv_end]:
                if el.find(f"./{W}pPr/{W}sectPr") is None:
                    body.remove(el)

    if not mode_cfg["include_declaration"]:
        children = list(body)
        decl_idx = next((i for i, el in enumerate(children) if _p_text(el).strip() == "学位论文独创性与使用授权声明"), None)
        if decl_idx is not None:
            for el in children[decl_idx:]:
                if el.tag == f"{W}p" and el.find(f"./{W}pPr/{W}sectPr") is None:
                    body.remove(el)

    for p in list(body.findall(f"{W}p")):
        txt = _p_text(p)
        t_el = p.find(f".//{W}t")
        if t_el is None:
            continue
        if title_en and txt.startswith("English title here"):
            t_el.text = title_en
        elif txt.startswith("Author Name"):
            if blind:
                t_el.text = f"Anonymous Author ({major or 'Major'})"
            elif author or major:
                t_el.text = f"{author or 'Author Name'} ({major or 'Major'})"
        elif txt.startswith("Directed by"):
            if blind:
                t_el.text = "Directed by 【匿名评审隐去】"
            elif advisor:
                t_el.text = f"Directed by {advisor}"
        elif txt.startswith("姓名："):
            if blind:
                t_el.text = "【匿名评审版本隐去姓名、性别、出生年月与籍贯】"
            elif author:
                t_el.text = txt.replace("【姓名】", author)
        elif txt.startswith("20XX年9月"):
            if blind:
                if "本科" in txt:
                    body.remove(p)
                else:
                    t_el.text = "【匿名评审版本隐去就读院校、培养单位与导师信息】"
            else:
                new_edu = txt
                if univ:
                    new_edu = new_edu.replace("【学校名称】", univ)
                if school:
                    new_edu = new_edu.replace("【培养单位】", school)
                if major:
                    new_edu = new_edu.replace("【专业名称】", major)
                t_el.text = new_edu
        elif blind and txt.startswith("1. Author, A."):
            t_el.text = "【匿名评审版本仅保留成果类型与作者排序，隐去个人、导师与合作者姓名】"
            first_run = next((r for r in p.findall(f"{W}r") if r.find(f".//{W}t") is t_el), None)
            if first_run is not None:
                for extra_r in p.findall(f"{W}r"):
                    if extra_r is not first_run:
                        p.remove(extra_r)
        elif blind and "【致谢内容" in txt:
            t_el.text = "【匿名评审版本隐去致谢，终稿恢复。】"
        elif blind and "修改自 Example et al." in txt:
            for t_node in p.findall(f".//{W}t"):
                if t_node.text and "修改自 Example et al." in t_node.text:
                    t_node.text = t_node.text.replace(
                        "修改自 Example et al.",
                        "盲审稿按学校要求标示协作贡献。修改自 Example et al.",
                    )
                    break


def _adapt_citation_style(body: ET.Element, citation_style: str) -> None:
    if citation_style != "gb7714":
        return

    # Convert Chapter 1 & Figure 1.1 superscript citations in placeholder scaffold to bracketed [1,2] / [1]
    for p in body.findall(f"{W}p"):
        txt = _p_text(p)
        if "示例引用详见图1.1与表1.1" in txt:
            for r in p.findall(f"{W}r"):
                rpr = r.find(f"{W}rPr")
                t = r.find(f"{W}t")
                if rpr is not None and t is not None and t.text == "1,2":
                    va = rpr.find(f"{W}vertAlign")
                    if va is not None:
                        rpr.remove(va)
                    t.text = "[1,2]"
        elif txt.startswith("图 1.1"):
            for r in p.findall(f"{W}r"):
                rpr = r.find(f"{W}rPr")
                t = r.find(f"{W}t")
                if rpr is not None and t is not None and t.text == "1":
                    va = rpr.find(f"{W}vertAlign")
                    if va is not None:
                        rpr.remove(va)
                    t.text = "[1]"

    children = list(body)
    ref_idx = next((i for i, el in enumerate(children) if _p_text(el).strip() == "参考文献"), None)
    cv_idx = next((i for i, el in enumerate(children) if _p_text(el).strip().startswith("个人简历")), None)
    if ref_idx is None or cv_idx is None or cv_idx <= ref_idx:
        return
    for el in children[ref_idx + 1 : cv_idx]:
        body.remove(el)
    ins_pos = list(body).index(next(el for el in list(body) if _p_text(el).strip().startswith("个人简历")))
    gb_entries = [
        _make_p(
            "[1] EXAMPLE A B, SAMPLE C D, PLACEHOLDER E F. A fictional example for document layout[J]. Fictional Journal of Examples, 2000, 1(1): 1-2.",
            "BibEntry",
        ),
        _make_p(
            "[2] 示例作者甲, 示例作者乙. 引文著录格式的虚构示例[J]. 虚构示例期刊, 2001, 2(1): 3-4.",
            "BibEntry",
        ),
    ]
    for off, node in enumerate(gb_entries):
        body.insert(ins_pos + off, node)


def _adapt_chapter_skeleton(body: ET.Element, mode_cfg: dict) -> None:
    profile = mode_cfg["chapter_profile"]
    self_term = mode_cfg.get("self_term", "本论文")

    children = list(body)
    idx_ch1_aim = next((i for i, el in enumerate(children) if _p_text(el).strip().startswith("1.2 ")), None)
    idx_ch2 = next((i for i, el in enumerate(children) if _p_text(el).strip().startswith("第二章")), None)
    idx_ref = next((i for i, el in enumerate(children) if _p_text(el).strip() == "参考文献"), None)

    if profile != "thesis" and idx_ch1_aim is not None:
        t_el = children[idx_ch1_aim].find(f".//{W}t")
        if t_el is not None:
            t_el.text = "1.2 本研究的目的、主要内容和意义" if profile == "proposal" else "1.2 本轮转课题的目的与内容"
        if idx_ch1_aim + 1 < len(children):
            for t_node in children[idx_ch1_aim + 1].iter(f"{W}t"):
                if t_node.text and "本论文" in t_node.text:
                    t_node.text = t_node.text.replace("本论文", self_term)

    if profile == "rotation":
        idx_ch1 = next((i for i, el in enumerate(children) if _p_text(el).strip().startswith("第一章")), None)
        if idx_ch1 is not None:
            t_el = children[idx_ch1].find(f".//{W}t")
            if t_el is not None:
                t_el.text = "第一章 轮转期间工作的背景和意义"

    if profile == "thesis" or idx_ch2 is None or idx_ref is None:
        return

    for el in children[idx_ch2:idx_ref]:
        body.remove(el)

    ref_node = next(el for el in list(body) if _p_text(el).strip() == "参考文献")
    ref_pos = list(body).index(ref_node)

    if profile == "proposal":
        new_nodes = [
            _make_p("第二章 实验材料与方法", "Heading1Chapter"),
            _make_p("2.1 实验材料", "Heading2Sec"),
            _make_p("【说明使用的资料、研究对象、试剂与仪器设备（型号与生产厂商间默认用中文逗号分隔，如 P-2000，Sutter Instrument）。】", "BodyTextZH"),
            _make_p("2.2 实验程序与数据分析", "Heading2Sec"),
            _make_p("【说明实验流程、分析方法及适用条件。实验程序中的时间单位须全文统一为中文或统一为英文缩写。】", "BodyTextZH"),
            _make_p("第三章 前期研究进展与初步讨论", "Heading1Chapter"),
            _make_p("3.1 【第一部分前期研究结果与讨论】", "Heading2Sec"),
            _make_p("【当前研究结果与讨论正文。】", "BodyTextZH"),
            _make_p("第四章 后续研究计划及方案设计", "Heading1Chapter"),
            _make_p("4.1 目标1：【核心研究目标】", "Heading2Sec"),
            _make_p("4.1.1 研究方案", "Heading3Subsec"),
            _make_p("【说明证据来源、研究设计及判断依据。】", "BodyTextZH"),
            _make_p("4.1.2 可行性分析", "Heading3Subsec"),
            _make_p("【结合前期基础与技术平台说明可行性。】", "BodyTextZH"),
            _make_p("4.1.3 预期结果", "Heading3Subsec"),
            _make_p("【说明预期成果及其判断标准。】", "BodyTextZH"),
            _make_p("4.1.4 可能的困难及应对方案", "Heading3Subsec"),
            _make_p("【潜在技术风险与替代路线。】", "BodyTextZH"),
        ]
    else:
        new_nodes = [
            _make_p("第二章 实验材料、方法和步骤", "Heading1Chapter"),
            _make_p("2.1 实验材料", "Heading2Sec"),
            _make_p("【说明使用的资料、仪器设备与研究对象（型号与厂商间默认用中文逗号分隔）。】", "BodyTextZH"),
            _make_p("2.2 实验程序和步骤", "Heading2Sec"),
            _make_p("【说明实施与分析步骤。实验程序中的时间单位须全文统一为中文或统一为英文缩写。】", "BodyTextZH"),
            _make_p("第三章 工作结果与讨论", "Heading1Chapter"),
            _make_p("3.1 【轮转主要结果与讨论】", "Heading2Sec"),
            _make_p("【说明轮转期间取得的结果以及仍待解决的问题。】", "BodyTextZH"),
            _make_p("第四章 轮转收获、不足与后续建议", "Heading1Chapter"),
            _make_p("4.1 方法与研究经验", "Heading2Sec"),
            _make_p("【轮转收获总结。】", "BodyTextZH"),
            _make_p("4.2 不足之处与后续建议", "Heading2Sec"),
            _make_p("【不足反思与课题推进建议。】", "BodyTextZH"),
        ]

    for off, node in enumerate(new_nodes):
        body.insert(ref_pos + off, node)


def _apply_page_geometry(doc_root: ET.Element, settings_root: ET.Element, page_cfg: dict) -> None:
    dims = _resolve_page_dxa(page_cfg)
    for sp in doc_root.iter(f"{W}sectPr"):
        pg_sz = sp.find(f"{W}pgSz")
        if pg_sz is not None:
            pg_sz.set(f"{W}w", str(dims["width_dxa"]))
            pg_sz.set(f"{W}h", str(dims["height_dxa"]))
            if dims['width_dxa'] > dims['height_dxa']:
                pg_sz.set(W + 'orient', 'landscape')
            else:
                pg_sz.attrib.pop(W + 'orient', None)
            if (dims['width_dxa'], dims['height_dxa']) != (11906, 16838):
                pg_sz.attrib.pop(W + 'code', None)
        pg_mar = sp.find(f"{W}pgMar")
        if pg_mar is not None:
            pg_mar.set(f"{W}top", str(dims["top"]))
            pg_mar.set(f"{W}bottom", str(dims["bottom"]))
            pg_mar.set(f"{W}left", str(dims["left"]))
            pg_mar.set(f"{W}right", str(dims["right"]))
            pg_mar.set(f"{W}header", str(dims["header"]))
            pg_mar.set(f"{W}footer", str(dims["footer"]))
            pg_mar.set(f"{W}gutter", str(dims["gutter"]))

    for tag, flag in (("mirrorMargins", page_cfg.get("mirror_margins", True)), ("evenAndOddHeaders", page_cfg.get("odd_even_headers", True))):
        node = settings_root.find(f"{W}{tag}")
        if flag and node is None:
            settings_root.insert(0, ET.Element(f"{W}{tag}"))
        elif not flag and node is not None:
            settings_root.remove(node)


def _apply_typography(items: dict[str, bytes], config: dict) -> None:
    if not config:
        return
    root = ET.fromstring(items['word/styles.xml'])
    groups = {
        'BodyTextZH': ('cjk_body_font', 'body_line_pt'),
        'AbstractBodyZH': ('cjk_body_font', None),
        'AbstractBodyEN': (None, None),
        'BibEntry': ('cjk_body_font', 'ref_line_pt'),
        'Heading1Chapter': ('cjk_heading_font', None),
        'Heading2Sec': ('cjk_heading_font', None),
        'Heading3Subsec': ('cjk_heading_font', None),
        'Heading4Para': ('cjk_heading_font', None),
    }
    for style in root.findall(W + 'style'):
        if style.get(W + 'styleId') not in groups:
            continue
        font_key, line_key = groups[style.get(W + 'styleId')]
        if line_key in config:
            line = round(20 * float(config[line_key]))
            if line <= 0:
                raise ValueError(f'{line_key} must be positive')
            spacing = style.find(f'./{W}pPr/{W}spacing')
            if spacing is not None:
                spacing.set(W + 'line', str(line))
                spacing.set(W + 'lineRule', 'exact')
        fonts = style.find(f'./{W}rPr/{W}rFonts')
        if fonts is None:
            continue
        if font_key in config:
            fonts.set(W + 'eastAsia', config[font_key])
            fonts.attrib.pop(W + 'eastAsiaTheme', None)
        if 'latin_font' in config:
            for key, theme in (('ascii', 'asciiTheme'), ('hAnsi', 'hAnsiTheme'), ('cs', 'cstheme')):
                fonts.set(W + key, config['latin_font'])
                fonts.attrib.pop(W + theme, None)
    items['word/styles.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _initialize_custom(args: argparse.Namespace, out_path: pathlib.Path) -> None:
    source = pathlib.Path(args.template).resolve()
    if same_path(source, out_path):
        raise ValueError('Output must not overwrite the user template')
    if args.logo:
        raise ValueError('Place the logo in the supplied Word template before using --template')
    items, package_type = read_template(source)
    fields = ('title_zh', 'title_en', 'author', 'student_id', 'advisor', 'major', 'date_zh', 'university', 'school')
    private = ('author', 'student_id', 'advisor')
    values = {key: getattr(args, key) for key in fields if getattr(args, key) is not None}
    if args.review_stage == 'blind':
        values.update(dict.fromkeys(private, '【匿名评审隐去】'))
    for value in values.values():
        if any(ord(char) < 32 for char in value):
            raise ValueError('Template field values must be single-line text without control characters')
    filled = fill_template(items, package_type, values)
    if args.review_stage == 'blind':
        tokens = [getattr(args, key) for key in private if getattr(args, key)]
        parsed = {name: ET.fromstring(data) for name, data in items.items() if name.endswith(('.xml', '.rels'))}
        leaked = identity_parts(parsed, tokens)
        if leaked:
            raise ValueError('Supplied identity text remains in template parts: ' + ', '.join(leaked))
        print('NOTICE: Blind mode fills identity placeholders only. Audit the completed document before sharing.')
    unused = sorted(key for key in values if key not in filled and getattr(args, key) is not None)
    if unused:
        print(f'NOTICE: No eligible placeholders for: {", ".join(unused)}')
    if args.profile_config or args.citation_style != 'auto':
        print('NOTICE: User template layout and citations are preserved. Use profile/citation options when auditing.')
    write_package(out_path, items, args.overwrite)
    print(f'CREATED={out_path}\nTEMPLATE={source}\nFILLED_FIELDS={filled.total()}')


def initialize_docx(args: argparse.Namespace) -> None:
    out_path = pathlib.Path(args.output).resolve()
    if out_path.suffix.lower() != '.docx':
        raise ValueError('Output must have a .docx extension')
    if same_path(out_path, MASTER_DOCX) or (args.profile_config and same_path(out_path, pathlib.Path(args.profile_config))):
        raise ValueError('Output must not overwrite an input file')
    if out_path.exists() and not args.overwrite:
        raise FileExistsError(f"Output exists (pass --overwrite to replace): {out_path}")
    if args.template:
        _initialize_custom(args, out_path)
        return

    preset = load_effective_preset(
        args.preset, args.profile_config, args.university, args.school, args.citation_style
    )
    mode_cfg = preset["modes"][args.mode]
    blind = args.review_stage == "blind"
    univ = preset.get("university_name", "【学校名称】")
    school = preset.get("school_name", "【培养单位】")
    header_text = mode_cfg["header_template"].format(university=univ, school=school)
    citation_style = preset.get("citation_style", "nature")

    logo_bytes: bytes | None = None
    logo_info: tuple[str, str, int, int] | None = None
    if args.logo:
        logo_path = pathlib.Path(args.logo).resolve()
        if same_path(out_path, logo_path):
            raise ValueError('Output must not overwrite the input logo')
        if not logo_path.exists():
            raise FileNotFoundError(f"--logo image not found: {logo_path}")
        ext, mime, px_w, px_h = _read_image_dims(logo_path)
        logo_bytes = logo_path.read_bytes()
        logo_info = ("rId99", ext, px_w, px_h)

    with zipfile.ZipFile(MASTER_DOCX, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    doc_root = ET.fromstring(items["word/document.xml"])
    settings_root = ET.fromstring(items["word/settings.xml"])
    body = doc_root.find(f"{W}body")
    if body is None:
        raise RuntimeError("Invalid master DOCX: missing w:body")

    _patch_cover(
        body,
        preset,
        mode_cfg,
        title_zh=args.title_zh,
        author=args.author,
        student_id=args.student_id,
        major=args.major,
        advisor=args.advisor,
        date_zh=args.date_zh,
        blind=blind,
        logo_info=logo_info,
    )
    _patch_front_and_back(
        body,
        mode_cfg,
        title_en=args.title_en,
        author=args.author,
        major=args.major,
        advisor=args.advisor,
        univ=univ,
        school=school,
        blind=blind,
    )
    _adapt_chapter_skeleton(body, mode_cfg)
    _adapt_citation_style(body, citation_style)
    _apply_page_geometry(doc_root, settings_root, preset.get("page", {}))
    if args.profile_config:
        overrides = json.loads(pathlib.Path(args.profile_config).read_text(encoding='utf-8'))
        _apply_typography(items, overrides.get('typography', {}))

    items["word/document.xml"] = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)
    items["word/settings.xml"] = ET.tostring(settings_root, encoding="utf-8", xml_declaration=True)

    for hdr_name in ("word/header1.xml", "word/header2.xml"):
        hdr_root = ET.fromstring(items[hdr_name])
        t_el = hdr_root.find(f".//{W}t")
        if t_el is not None:
            t_el.text = header_text
        items[hdr_name] = ET.tostring(hdr_root, encoding="utf-8", xml_declaration=True)

    if logo_bytes is not None and logo_info is not None:
        rel_id, ext, _, _ = logo_info
        mime = "image/png" if ext == "png" else "image/jpeg"
        ET.register_namespace("", CT_NS)
        ct_root = ET.fromstring(items["[Content_Types].xml"])
        existing_exts = {d.get("Extension", "").lower() for d in ct_root.findall(f"{{{CT_NS}}}Default")}
        if ext.lower() not in existing_exts:
            ET.SubElement(ct_root, f"{{{CT_NS}}}Default", {"Extension": ext, "ContentType": mime})
        items["[Content_Types].xml"] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

        ET.register_namespace("", REL_NS)
        rels_root = ET.fromstring(items["word/_rels/document.xml.rels"])
        ET.SubElement(
            rels_root,
            f"{{{REL_NS}}}Relationship",
            {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": f"media/logo.{ext}",
            },
        )
        items["word/_rels/document.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
        items[f"word/media/logo.{ext}"] = logo_bytes

    ET.register_namespace("cp", CP_NS)
    core_root = ET.fromstring(items["docProps/core.xml"])
    title_el = core_root.find(f"{{{DC_NS}}}title")
    if title_el is not None:
        title_el.text = f"{univ} {args.mode} ({args.review_stage})"
    items["docProps/core.xml"] = ET.tostring(core_root, encoding="utf-8", xml_declaration=True)

    write_package(out_path, items, args.overwrite)

    print(f"CREATED={out_path}")
    print(f"PRESET={preset['preset_name']} CITATION_STYLE={citation_style}")
    print(f"MODE={args.mode} REVIEW_STAGE={args.review_stage}")


def add_arguments(p: argparse.ArgumentParser) -> None:
    p.add_argument("--output", required=True, help="Output .docx path")
    p.add_argument("--template", help="Use a user-supplied .docx or .dotx as the layout and content source")
    p.add_argument("--mode", choices=("phd-thesis", "master-thesis", "proposal", "midterm", "rotation"), default="phd-thesis")
    p.add_argument("--review-stage", choices=("final", "blind"), default="final")
    p.add_argument("--preset", choices=("cn-lifesci-default", "cn-gb7714-general"), default="cn-lifesci-default")
    p.add_argument("--profile-config", help="Optional custom university/department JSON override file")
    p.add_argument("--citation-style", choices=("auto", "nature", "gb7714"), default="auto")
    p.add_argument("--university", help="Override university name on cover and headers")
    p.add_argument("--school", help="Override school/department name on cover and headers")
    p.add_argument("--title-zh", help="Chinese thesis/report title")
    p.add_argument("--title-en", help="English thesis/report title")
    p.add_argument("--author", help="Author name (redacted automatically in blind mode)")
    p.add_argument("--student-id", help="Student ID (redacted automatically in blind mode)")
    p.add_argument("--major", help="Major / academic discipline")
    p.add_argument("--advisor", help="Advisor name and title (redacted in blind mode)")
    p.add_argument("--date-zh", help="Cover date line, e.g. '二〇二六年五月'")
    p.add_argument("--logo", help="Optional PNG/JPEG university emblem to embed at top of cover")
    p.add_argument("--overwrite", action="store_true", help="Overwrite existing output file")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Initialize a bio-thesis DOCX from the clean neutral master template.")
    add_arguments(parser)
    args = parser.parse_args(argv)
    try:
        initialize_docx(args)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
