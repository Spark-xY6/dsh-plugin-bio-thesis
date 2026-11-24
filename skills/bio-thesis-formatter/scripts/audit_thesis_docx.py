#!/usr/bin/env python3
"""Read-only OOXML & 67-item compliance auditor using Python stdlib only."""

from __future__ import annotations

import argparse
import hashlib
import json
import pathlib
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import asdict, dataclass

from docx_package import has_main_relationship, identity_parts, read_template, resolve_target, same_path, staged_file, word_main_type
from init_thesis_docx import _resolve_page_dxa, load_effective_preset
from repair_margins import _direct_equal_length_compatible, align_section_indices, elem_has_field

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SVG_NS = "http://www.w3.org/2000/svg"
W = f"{{{W_NS}}}"
A = f"{{{A_NS}}}"
R = f"{{{R_NS}}}"

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent.resolve()
SCHOOL_CHECKS = {
    'Err-9', 'Layout-9', 'Err-8', 'Layout-BodySpacing', 'Ref-9', 'Tbl-3Line',
    'Fig-4', 'Fig-5', 'Fig-9-10', 'Acad-2', 'Acad-14', 'Acad-CV',
    'Err-2-11-12', 'Err-7', 'Err-14', 'Layout-7',
}

CH1_STYLES = {"Heading1Chapter", "04", "Heading1", "heading 1"}
H2_STYLES = {"Heading2Sec", "05XX", "Heading2", "heading 2"}
H3_STYLES = {"Heading3Subsec", "06XXX", "Heading3", "heading 3"}
H4_STYLES = {"Heading4Para", "07XXXX", "Heading4", "heading 4"}
ABS_ZH_STYLES = {"AbstractBodyZH", "02"}
ABS_EN_STYLES = {"AbstractBodyEN", "03"}
BODY_STYLES = {"BodyTextZH", "08", "Normal", "a", "BodyText"}
FIG_CAP_STYLES = {"FigureCaption", "09"}
TBL_TITLE_STYLES = {"TableTitle", "100"}
TBL_NOTE_STYLES = {"TableNote", "12"}
BIB_STYLES = {"BibEntry", "13ref"}

ENZYME_NO_SPACE_RE = re.compile(
    r"\b(Taq|EcoR|BamH|Hind|Xho|Not|Nde|Kpn|Sal|Spe|Xba|Pst|Sma|Bgl|Nco|Sac|Apa|Sca|Mlu|Nhe|Age|Bsa|BsmB)(I{1,3}|IV|V|VI{0,3})\b"
)
BUILTIN_SPECIES = (
    "Homo sapiens", "Mus musculus", "Drosophila melanogaster", "Drosophila simulans",
    "Caenorhabditis elegans", "Escherichia coli", "Saccharomyces cerevisiae",
    "Danio rerio", "Arabidopsis thaliana", "Rattus norvegicus",
)
BAD_UNIT_U_RE = re.compile(r"(?<![A-Za-z0-9_])(\d+(?:\.\d+)?\s*)?u([LlgM])\b")
NO_SPACE_UNIT_RE = re.compile(
    r"(?<![A-Za-z0-9\-–—_])(\d+(?:\.\d+)?)(mL|µL|μL|nL|mg|µg|μg|ng|kg|mM|µM|μM|nM|pM|kb|Mb|kDa|Da|rpm)\b"
)
CJK_TIME_RE = re.compile(
    r"(?:(?:\d+(?:\.\d+)?|[一二三四五六七八九十两半]+)\s*(?:个?\s*小时|分钟|秒钟?)|\d+(?:\.\d+)?\s*[周天]|[一二三五六七八九十两]+\s*[周天]|四周(?:后|内|前|龄|期))"
)
ENG_TIME_WORD_RE = re.compile(
    r"(?<![.\dA-Za-z])\d+(?:\.\d+)?\s*(?:hr|hrs|hour|hours|min|mins|minute|minutes|sec|secs|second|seconds|wk|wks|week|weeks)\b",
    re.IGNORECASE,
)
ENG_TIME_SYM_RE = re.compile(
    r"(?<![.\dA-Za-z])(?!19\d0s\b|20\d0s\b|1[68]s\b|2[38]s\b|\d+s(?=\s*(?:核糖体|亚基|rRNA|ribosom|subunit|年代)))\d+(?:\.\d+)?\s*(?:h|s|d)\b(?!\s*(?:电泳|培养|重建|结构|图|gel))"
)
FIG_TBL_REF_STRIP_RE = re.compile(r"(?:Fig\.?|Figure|Panel|Table|图|表)\s*[A-Z]?\d+(?:\.\d+)?\s*[A-Za-z]*")
FORMULA_RE = re.compile(r'(?<![A-Za-z0-9_])(?:H2O|CO2|Ca2\+|Mg2\+|O2|N2)(?![A-Za-z0-9_])')
SUBPANEL_BAD_SEP_RE = re.compile(r"(?:补充图|图)\s*[A-D1-9]\d*[.\-]\d+\s*[A-Z](?:\s*[,、]\s*[A-Z])+")
CJK_ASCII_PUNCT_RE = re.compile(r"[\u4e00-\u9fff][,;:\(\)][\u4e00-\u9fff]")
VENDOR_KEYWORDS_RE = re.compile(
    r"(?:Instrument|Instruments|Bio-Rad|Thermo|Fisher|Sigma|Aldrich|Abcam|Illumina|Beckman|Olympus|Zeiss|Leica|Nikon|Corning|Gibco|Invitrogen|Takara|Vazyme|Beyotime|CST|NEB|Qiagen|Roche|Merck|Agilent|Waters|Bruker|Sutter|Eppendorf|BGI|Novogene|Technologies|Scientific|Biosciences|Co\.|Inc\.|Ltd\.|Corp\.|GmbH|公司|生物|科技|仪器|试剂|集团)",
    re.IGNORECASE,
)
VENDOR_BRAND_RE = re.compile(
    r"(?:Instrument|Instruments|Bio-Rad|Thermo|Fisher|Sigma|Aldrich|Abcam|Illumina|Beckman|Olympus|Zeiss|Leica|Nikon|Corning|Gibco|Invitrogen|Takara|Vazyme|Beyotime|CST|NEB|Qiagen|Roche|Merck|Agilent|Waters|Bruker|Sutter|Eppendorf|BGI|Novogene|Technologies|Scientific|Biosciences|Co\.|Inc\.|Ltd\.|Corp\.|GmbH|有限公司|公司|集团|生物科技|股份)",
    re.IGNORECASE,
)
PAREN_ZH_CONTENT_RE = re.compile(r"（([^（）\n]{1,100})）")
PAREN_VENDOR_DUNHAO_RE = re.compile(
    r"[（(]([^（）()\n]{1,70}、[^（）()\n]{1,70})[）)]"
)
CJK_HALF_SPACE_RE = re.compile(r"[\u4e00-\u9fff][ \t]+[\u4e00-\u9fff]")
ASCII_FULL_SPACE_RE = re.compile(r"[A-Za-z0-9.,;:!?%°]\u3000+")
FIRST_PERSON_RE = re.compile(
    r"(?:(?<![自本超忘大小敌物你他非真])我们?(?!国)|(?<![A-Za-z样日基技版剧脚成原书话读拓复副正孤珍绝])本人(?![口数群员均士籍类物])|(?<![A-Za-z])(?:笔者|咱们))"
)
DOI_RE = re.compile(r"(?:\bdoi\s*:\s*|https?://(?:dx\.)?doi\.org/|\b10\.\d{4,9}/\S+)", re.IGNORECASE)
AHEAD_OF_PRINT_RE = re.compile(
    r"(?:advance online publication|online ahead of print|epub ahead of print|ahead of print|in press|preprint|biorxiv|medrxiv|arxiv|research square|ssrn|在线发表|网络首发|预印本|待刊)",
    re.IGNORECASE,
)
PUBLISHED_VOL_PAGE_RE = re.compile(r"\b\d+\s*[,:]\s*[A-Za-z]?\d+(?:\s*[–-]\s*[A-Za-z]?\d+)?\s*\(\d{4}\)")
ABBR_PAREN_RE = re.compile(r"([^（(]{2,45}?)[（(](?:([^（）()]+?)[，,]\s*)?([A-Za-z][A-Za-z0-9\-/]{1,14})[）)]")

LOWERCASE_ALLOWED_JOURNALS = {"elife", "iscience", "biorxiv", "medrxiv", "plos"}
UPPERCASE_EXACT_TOKENS = {"RNA", "DNA", "mRNA", "tRNA", "rRNA", "lncRNA", "miRNA", "siRNA", "XCI", "CRISPR", "Hi-C", "ATAC", "ChIP", "SNP", "QTL", "3D"}


@dataclass
class CheckResult:
    check_id: str
    checklist_ref: str
    status: str  # PASS | FAIL | WARN | INFO
    summary: str
    details: list[str]


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def _p_text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(f"{W}t"))


def _live_p_text(el: ET.Element) -> str:
    """Extract live run text inside a paragraph while excluding deleted runs and nested textboxes."""
    return "".join(_r_text(r) for r in _live_word_items(el, W + "r"))


def _r_text(r: ET.Element) -> str:
    return "".join(t.text or "" for t in r.iter(f"{W}t"))


def _p_style(p: ET.Element) -> str | None:
    ps = p.find(f"./{W}pPr/{W}pStyle")
    return ps.get(f"{W}val") if ps is not None else None


def _tag_local(el: ET.Element) -> str:
    return el.tag.split("}", 1)[-1]


def _on(node: ET.Element | None) -> bool:
    return node is not None and node.get(W + 'val', '1') not in ('0', 'false', 'off')


def _fmt_dxa(value: int | str | None) -> str:
    if isinstance(value, int):
        return f"{value * 2.54 / 1440:.2f} cm ({value} dxa)"
    return str(value)


def _iter_story_blocks(container: ET.Element):
    """Yield p and tbl blocks in document order, unwrapping w:sdt controls while skipping deleted/textbox content."""
    for child in list(container):
        if child.tag in (W + 'del', W + 'moveFrom', W + 'txbxContent'):
            continue
        if child.tag == W + 'sdt':
            content = child.find(W + 'sdtContent')
            if content is not None:
                yield from _iter_story_blocks(content)
        elif child.tag in (W + 'p', W + 'tbl'):
            yield child


def _load_theme_fonts(theme_root: ET.Element | None) -> dict[str, str]:
    fonts: dict[str, str] = {}
    if theme_root is None:
        return fonts
    for role, tag in (('major', 'majorFont'), ('minor', 'minorFont')):
        node = theme_root.find(f'.//{A}{tag}')
        if node is None:
            continue
        latin = node.find(A + 'latin')
        if latin is not None and latin.get('typeface'):
            fonts[f'{role}HAnsi'] = latin.get('typeface', '')
            fonts[f'{role}Ascii'] = latin.get('typeface', '')
        ea = node.find(A + 'ea')
        ea_face = ea.get('typeface', '') if ea is not None else ''
        if not ea_face:
            for alt in node.findall(A + 'font'):
                if alt.get('script') in ('Hans', 'Hant', 'EastAsia') and alt.get('typeface'):
                    ea_face = alt.get('typeface', '')
                    break
        if ea_face:
            fonts[f'{role}EastAsia'] = ea_face
    return fonts


def _load_numbering(numbering_root: ET.Element | None) -> tuple[dict[str, str], dict[tuple[str, str], tuple[str, str]]]:
    """Return (numId->abstractNumId, (abstractNumId, ilvl)->(lvlText, suff))."""
    num_to_abs: dict[str, str] = {}
    levels: dict[tuple[str, str], tuple[str, str]] = {}
    if numbering_root is None:
        return num_to_abs, levels
    for abs_el in numbering_root.findall(W + 'abstractNum'):
        abs_id = abs_el.get(W + 'abstractNumId')
        if not abs_id:
            continue
        for lvl in abs_el.findall(W + 'lvl'):
            ilvl = lvl.get(W + 'ilvl', '0')
            lt = lvl.find(W + 'lvlText')
            sf = lvl.find(W + 'suff')
            levels[(abs_id, ilvl)] = (
                lt.get(W + 'val', '') if lt is not None else '',
                sf.get(W + 'val', 'tab') if sf is not None else 'tab',
            )
    for num_el in numbering_root.findall(W + 'num'):
        nid = num_el.get(W + 'numId')
        aid = num_el.find(W + 'abstractNumId')
        if nid and aid is not None and aid.get(W + 'val'):
            num_to_abs[nid] = aid.get(W + 'val', '')
    return num_to_abs, levels


def _section_geometry(section: ET.Element) -> dict[str, int | str | None]:
    geometry = {}
    for tag, keys in (('pgSz', ('w', 'h')), ('pgMar', ('top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'))):
        node = section.find(W + tag)
        for key in keys:
            value = node.get(W + key) if node is not None else None
            geometry[key] = int(value) if value is not None else None
        if tag == 'pgSz':
            if node is not None:
                orient = node.get(W + 'orient')
                if not orient and geometry.get('w') and geometry.get('h') and geometry['w'] > geometry['h']:
                    orient = 'landscape'
                geometry['orient'] = orient or 'portrait'
            else:
                geometry['orient'] = None
    return geometry


def _instructions(root: ET.Element, command: str) -> bool:
    return elem_has_field(root, command)


def _is_toc_para(el: ET.Element) -> bool:
    ps = (_p_style(el) or '').lower()
    return ps.startswith('toc') or _instructions(el, 'PAGEREF') or _instructions(el, 'TOC')


def _live_sections(doc_root: ET.Element) -> list[ET.Element]:
    body = doc_root.find(W + 'body')
    return list(_live_word_items(body, W + 'sectPr')) if body is not None else []


def _section_numbering(sections, parsed, rels_map):
    result = []
    headers = {}
    footers = {}
    for section in sections:
        numbering = section.find(W + 'pgNumType')
        fmt = numbering.get(W + 'fmt', 'decimal') if numbering is not None else 'decimal'
        start = numbering.get(W + 'start') if numbering is not None else None
        section_type = section.find(W + 'type')
        kind = section_type.get(W + 'val', 'nextPage') if section_type is not None else 'nextPage'
        for tag, inherited in (('headerReference', headers), ('footerReference', footers)):
            for reference in section.findall(W + tag):
                part = rels_map.get(reference.get(R + 'id', ''), '')
                inherited[reference.get(W + 'type', 'default')] = parsed.get(part)
        active = tuple(sorted(role for role, root in footers.items()
                              if root is not None and _instructions(root, 'PAGE')))
        header_roles = tuple(sorted(role for role, root in headers.items() if root is not None))
        result.append((fmt, start, kind, _on(section.find(W + 'titlePg')), active, header_roles))
    return result


def _section_meta(geometry_list, numbering_list) -> list[dict]:
    meta = []
    for g, (fmt, start, kind, title_pg, active, _) in zip(geometry_list, numbering_list):
        orient = g.get('orient') or ('landscape' if (g.get('w') or 0) > (g.get('h') or 0) else 'portrait')
        meta.append({
            'orient': orient,
            'fmt': fmt,
            'start': start,
            'kind': kind,
            'title_pg': title_pg,
            'has_page': bool(active),
        })
    return meta


def _template_snapshot(template_path: str) -> dict:
    items, _ = read_template(pathlib.Path(template_path))
    parts = {k: ET.fromstring(v) for k, v in items.items() if k.endswith(('.xml', '.rels'))}
    document = parts['word/document.xml']
    rels = parts.get('word/_rels/document.xml.rels')
    rels_map = {r.get('Id', ''): resolve_target('word/document.xml', r.get('Target', ''))
                for r in rels if r.get('TargetMode', 'Internal') == 'Internal'} if rels is not None else {}
    sections = _live_sections(document)
    geom = [_section_geometry(sec) for sec in sections]
    num = _section_numbering(sections, parts, rels_map)
    return {
        'settings': parts.get('word/settings.xml'),
        'geometry': geom,
        'numbering': num,
        'section_meta': _section_meta(geom, num),
        'toc': _instructions(document, 'TOC'),
        'style_props': _style_properties(parts.get('word/styles.xml')),
    }


def _style_properties(root):
    def fingerprint(node):
        return (node.tag, tuple(sorted(node.attrib.items())), tuple(sorted(fingerprint(child) for child in node)))

    props = {}
    if root is not None:
        for style in root.findall(W + 'style'):
            key = style.get(W + 'styleId')
            props[key] = tuple(sorted(fingerprint(child) for child in style
                                      if _tag_local(child) in ('basedOn', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr')))
        defaults = root.find(W + 'docDefaults')
        if defaults is not None:
            props['<docDefaults>'] = fingerprint(defaults)
    return props


def _use_audit_policy(results, preset, template, explicit_citation):
    policy = preset.get('audit', {})
    configured = {}
    for key in ('disabled_checks', 'template_checks'):
        ids = policy.get(key, [])
        if not isinstance(ids, list) or any(not isinstance(item, str) or item not in SCHOOL_CHECKS for item in ids):
            raise ValueError(f'audit.{key} must list supported school-specific check IDs')
        configured[key] = set(ids)
    for check in results:
        key = check.check_id
        disabled = key in configured['disabled_checks']
        differs_by_school = template and key in SCHOOL_CHECKS and key not in configured['template_checks']
        unknown_citations = template and not explicit_citation and key in {'Ref-1', 'Ref-2-7', 'Ref-8'}
        if disabled or differs_by_school or unknown_citations:
            check.status = 'INFO'
            check.summary = ('模板未指定引文体系，请用 --citation-style 或配置指定' if unknown_citations
                             else '此校级规则未启用，按本校规范复核')
            check.details = []


DIRECT_PROPERTIES = {
    'font_size': (W + 'sz', W + 'szCs'),
    'font_face': (W + 'rFonts',),
    'color': (W + 'color',),
}
FONT_ASSIGNMENTS = {W + name for name in ('ascii', 'hAnsi', 'eastAsia', 'cs', 'asciiTheme',
                                           'hAnsiTheme', 'eastAsiaTheme', 'cstheme', 'csTheme')}


def _live_word_items(parent, tag):
    """Yield top-level matches in main-story order, omitting deleted content and text boxes."""
    stack = list(reversed(parent))
    while stack:
        item = stack.pop()
        if item.tag in (W + 'del', W + 'moveFrom', W + 'txbxContent', W + 'sectPrChange'):
            continue
        if item.tag == W + 'tr' and item.find(f'./{W}trPr/{W}del') is not None:
            continue
        if item.tag == W + 'tc' and item.find(f'./{W}tcPr/{W}cellDel') is not None:
            continue
        if item.tag == tag:
            yield item
        else:
            stack.extend(reversed(item))


def _formula_review(paragraph, styles):
    text = []
    alignment = []
    p_chain = styles._chain(_p_style(paragraph) or styles.default_paragraph_style)
    for run in _live_word_items(paragraph, W + 'r'):
        value = ''.join(t.text or '' for t in run.findall(W + 't'))
        if not value:
            continue
        char_style = run.find(f'./{W}rPr/{W}rStyle')
        r_chain = styles._chain(char_style.get(W + 'val')) if char_style is not None else []
        align = styles.default_vertical_alignment
        for node in [*p_chain, *r_chain, run]:
            override = node.find(f'./{W}rPr/{W}vertAlign')
            if override is not None:
                align = override.get(W + 'val')
        text.append(value)
        alignment.extend([align] * len(value))
    issues = []
    for match in FORMULA_RE.finditer(''.join(text)):
        formula = match.group()
        first = match.start() + formula.index('2')
        count = 2 if formula.endswith('+') else 1
        expected = 'superscript' if count == 2 else 'subscript'
        if any(value != expected for value in alignment[first:first + count]):
            issues.append(f'请核对化学式 {formula} 的数字或电荷是否按要求排为上下标')
    return issues


def _parse_svg_font_pt(raw_val: str | None) -> float | None:
    if not raw_val:
        return None
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*(pt|px)?\s*$", raw_val, re.IGNORECASE)
    if not m:
        return None
    val = float(m.group(1))
    # Bare unitless SVG font-size values represent px/user-units (96 dpi -> 0.75 pt).
    unit = (m.group(2) or "px").lower()
    return val * 0.75 if unit == "px" else val


def _extract_svg_node_pt(node: ET.Element) -> float | None:
    pt = _parse_svg_font_pt(node.get("font-size"))
    if pt is None:
        style_attr = node.get("style") or ""
        ms = re.search(r"font-size\s*:\s*([^;]+)", style_attr, re.IGNORECASE)
        if ms:
            pt = _parse_svg_font_pt(ms.group(1))
    return pt


def _walk_svg_text_sizes(node: ET.Element, inherited_pt: float | None, part: str, min_fig_pt: float, issues: list[str]) -> int:
    checked = 0
    own_pt = _extract_svg_node_pt(node)
    eff_pt = own_pt if own_pt is not None else inherited_pt
    local = _tag_local(node)
    if local in ("text", "tspan"):
        own_txt = (node.text or "").strip()
        if own_txt and eff_pt is not None:
            checked += 1
            if eff_pt < min_fig_pt - 1e-6:
                issues.append(f"{part} 内文字小于 6 号字 ({eff_pt:.1f} < {min_fig_pt:g} pt): {own_txt[:24]!r}")
    for child in node:
        checked += _walk_svg_text_sizes(child, eff_pt, part, min_fig_pt, issues)
        if local in ("text", "tspan") and (child.tail or "").strip() and eff_pt is not None:
            checked += 1
            if eff_pt < min_fig_pt - 1e-6:
                issues.append(f"{part} 内文字小于 6 号字 ({eff_pt:.1f} < {min_fig_pt:g} pt): {(child.tail or '').strip()[:24]!r}")
    return checked


def _inspect_figure_font_sizes(
    main_blocks: list[ET.Element],
    zf: zipfile.ZipFile,
    names: set[str],
    rels_map: dict[str, str],
    min_fig_pt: float,
) -> tuple[int, list[str]]:
    checked = 0
    issues: list[str] = []
    seen_svgs: set[str] = set()
    for blk in main_blocks:
        for box in [*blk.iter(f"{W}drawing"), *blk.iter(f"{W}pict")]:
            for ar in box.iter(f"{A}r"):
                txt = "".join(t.text or "" for t in ar.iter(f"{A}t")).strip()
                if not txt:
                    continue
                rpr = ar.find(f"./{A}rPr")
                sz = rpr.get("sz") if rpr is not None else None
                if sz and sz.isdigit():
                    checked += 1
                    pt = int(sz) / 100.0
                    if pt < min_fig_pt - 1e-6:
                        issues.append(f"图中 DrawingML 文字小于 6 号字 ({pt:.1f} < {min_fig_pt:g} pt): {txt[:24]!r}")
            for wr in box.iter(f"{W}r"):
                txt = "".join(t.text or "" for t in wr.iter(f"{W}t")).strip()
                if not txt:
                    continue
                sz_el = wr.find(f"./{W}rPr/{W}sz")
                sz_v = sz_el.get(W + "val") if sz_el is not None else None
                if sz_v and sz_v.isdigit():
                    checked += 1
                    pt = int(sz_v) / 2.0
                    if pt < min_fig_pt - 1e-6:
                        issues.append(f"图中文本框文字小于 6 号字 ({pt:.1f} < {min_fig_pt:g} pt): {txt[:24]!r}")
            for el in box.iter():
                rid = el.get(f"{R}embed")
                if not rid:
                    continue
                part = rels_map.get(rid, "")
                if part.lower().endswith(".svg") and part in names and part not in seen_svgs:
                    seen_svgs.add(part)
                    try:
                        svg_root = ET.fromstring(zf.read(part))
                    except (ET.ParseError, UnicodeDecodeError) as exc:
                        checked += 1
                        issues.append(f"{part} 内嵌 SVG 解析失败: {exc}")
                        continue
                    checked += _walk_svg_text_sizes(svg_root, None, part, min_fig_pt, issues)
    return checked, issues


def _direct_format_checks(policy, styles, body):
    if not isinstance(policy, dict):
        raise ValueError('audit must be a JSON object')
    supported = {'template_checks', 'disabled_checks', 'allow_live_citations', 'forbid_published_doi', 'direct_formatting'}
    unknown = policy.keys() - supported
    if unknown:
        raise ValueError(f'Unknown audit option(s): {", ".join(sorted(unknown))}. '
                         f'Supported: {", ".join(sorted(supported))}.')
    for bool_key in ('allow_live_citations', 'forbid_published_doi'):
        if bool_key in policy and not isinstance(policy[bool_key], bool):
            raise ValueError(f'audit.{bool_key} must be true or false')
    rules = policy.get('direct_formatting', [])
    if not isinstance(rules, list):
        raise ValueError('audit.direct_formatting must be a list of rules')
    if not rules:
        return []
    paragraph_styles = {sid for sid, st in styles.styles.items() if st.get(W + 'type') == 'paragraph'}
    seen = set()
    results = []
    for index, rule in enumerate(rules):
        where = f'audit.direct_formatting[{index}]'
        if not isinstance(rule, dict) or set(rule) != {'id', 'paragraph_style', 'forbid'}:
            raise ValueError(f'{where} requires exactly id, paragraph_style and forbid')
        rule_id, sid, forbid = rule['id'], rule['paragraph_style'], rule['forbid']
        if not isinstance(rule_id, str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._-]*', rule_id) or rule_id in seen:
            raise ValueError(f'{where}.id must be unique and match [A-Za-z0-9][A-Za-z0-9._-]*')
        seen.add(rule_id)
        if not isinstance(sid, str) or sid not in paragraph_styles:
            raise ValueError(f'{where}.paragraph_style {sid!r} is not a paragraph style ID in word/styles.xml')
        if (not isinstance(forbid, list) or not forbid
                or any(not isinstance(key, str) or key not in DIRECT_PROPERTIES for key in forbid)
                or len(set(forbid)) != len(forbid)):
            raise ValueError(f'{where}.forbid must list distinct properties from {", ".join(DIRECT_PROPERTIES)}')

        applicable = {style_id for style_id in paragraph_styles if styles.styles[sid] in styles._chain(style_id)}
        checked = 0
        details = []
        for number, paragraph in enumerate(_live_word_items(body, W + 'p'), 1):
            if (_p_style(paragraph) or styles.default_paragraph_style) not in applicable:
                continue
            checked += 1
            found = set()
            for run in _live_word_items(paragraph, W + 'r'):
                props = run.find(W + 'rPr')
                if props is not None and any(t.text for t in run.findall(W + 't')):
                    for prop in props:
                        # A script-selection hint alone does not assign a font family.
                        if prop.tag != W + 'rFonts' or FONT_ASSIGNMENTS.intersection(prop.attrib):
                            found.add(prop.tag)
            violations = [key for key in forbid if found.intersection(DIRECT_PROPERTIES[key])]
            if violations:
                details.append(f'paragraph={number} properties={",".join(violations)}')
        status = 'FAIL' if details else 'PASS' if checked else 'WARN'
        summary = (f'样式 {sid} 及其派生样式共匹配 {checked} 段，检出直接格式 {len(details)} 段'
                   if checked else f'正文中无段落使用 {sid} 或其派生样式，请核对所选样式与检查范围')
        results.append(CheckResult('DirectFormat-' + rule_id, '自定义段落样式规则', status, summary, details))
    return results


def _border_visible(border) -> bool:
    return border is not None and border.get(W + 'val', 'nil') not in ('nil', 'none')


def _matches_rule(border, sizes: str | tuple[str, ...]) -> bool:
    allowed = (sizes,) if isinstance(sizes, str) else sizes
    return border is not None and border.get(W + 'val') == 'single' and border.get(W + 'sz') in allowed


def _table_problems(table: ET.Element) -> tuple[list[str], list[str], bool]:
    props = table.find(W + 'tblPr')
    borders = table.find(f'./{W}tblPr/{W}tblBorders')
    if borders is None:
        inherited = props is not None and props.find(W + 'tblStyle') is not None
        if inherited:
            return [], ['无显式边框，须验证表格样式解析结果'], True
        return ['缺少显式表格三线边框'], [], False
    problems: list[str] = []
    warnings: list[str] = []
    for side in ('top', 'bottom'):
        if not _matches_rule(borders.find(W + side), ('12', '8')):
            problems.append(f'{side} 应为 1.0–1.5 pt 单线')
    for side in ('left', 'right', 'start', 'end', 'insideV', 'insideH'):
        if _border_visible(borders.find(W + side)):
            problems.append(f'{side} 存在额外边框')
    rows = table.findall(W + 'tr')
    if not rows:
        return problems + ['表格缺少可核验的行'], warnings, False
    header_count = 0
    for row in rows:
        if not _on(row.find(f'./{W}trPr/{W}tblHeader')):
            break
        header_count += 1
    if not header_count:
        if len(rows) > 5:
            problems.append('首行未标记为可重复表头')
        else:
            warnings.append('首行未标记为可重复表头（跨页时建议勾选“在各页顶端重复出现”）')
        header_count = 1
    for index, row in enumerate(rows):
        if row.find(f'./{W}tblPrEx/{W}tblBorders') is not None:
            problems.append(f'第 {index + 1} 行存在表级边框例外，须人工确认')
        for cell in row.findall(W + 'tc'):
            cell_borders = cell.find(f'./{W}tcPr/{W}tcBorders')
            span_el = cell.find(f'./{W}tcPr/{W}gridSpan')
            span_raw = span_el.get(W + 'val', '1') if span_el is not None else '1'
            has_span = span_raw.isdigit() and int(span_raw) > 1
            if index == header_count - 1:
                bottom = cell_borders.find(W + 'bottom') if cell_borders is not None else None
                if not _matches_rule(bottom, ('6', '4')):
                    problems.append('表头分隔线应为 0.5–0.75 pt 单线')
            if cell_borders is None:
                continue
            for border in cell_borders:
                if not _border_visible(border):
                    continue
                side = _tag_local(border)
                allowed = (
                    side == 'bottom' and index == header_count - 1
                    or side == 'bottom' and index < header_count - 1 and has_span and _matches_rule(border, ('6', '4'))
                    or side == 'top' and index == 0 and _matches_rule(border, ('12', '8'))
                    or side == 'bottom' and index == len(rows) - 1 and _matches_rule(border, ('12', '8'))
                    or side == 'top' and index == header_count and _matches_rule(border, ('6', '4'))
                )
                if not allowed:
                    problems.append(f'第 {index + 1} 行单元格存在额外 {side} 边框')
    return list(dict.fromkeys(problems)), list(dict.fromkeys(warnings)), False


def _nature_author_title(runs, paragraph, styles) -> list[str]:
    head = []
    for run in runs:
        if styles.effective_italic(run, paragraph) and _r_text(run).strip():
            break
        head.append(_r_text(run))
    prefix = ''.join(head).strip()
    issues = []
    if re.search(r'(?:^|[,&]\s*)[A-Z][\w’\'-]+,\s*(?:[A-Z]{2,4}\.|[A-Z]\.\s*[A-Z]\.(?:\s*[A-Z]\.)?)', prefix):
        if re.search(r',\s*[A-Z]{2,4}\.|\b[A-Z]\.[A-Z]\.', prefix):
            issues.append('作者姓名缩写之间缺少空格')
    if re.search(r'\b[A-Z]\.\s+and\s+[A-Z][\w’\'-]+,', prefix):
        issues.append('作者之间使用 and，应为 &')
    title = re.search(r'(?:\.\s+)([^.!?]+)[.!?]$', prefix)
    if title is not None:
        text = title.group(1).strip()
        if len(re.findall(r'[A-Za-z]{3,}', text)) >= 3 and text.isupper():
            issues.append('英文论文题名为全大写，须核对句式大小写')
    return issues


class StyleResolver:
    def __init__(self, styles_root: ET.Element | None, theme_root: ET.Element | None = None) -> None:
        self.styles: dict[str, ET.Element] = {}
        self.default_paragraph_style: str | None = None
        self.default_vertical_alignment: str | None = None
        self.theme_fonts = _load_theme_fonts(theme_root)
        self.default_fonts: dict[str, str] = {}
        self.default_p_jc: str = "left"
        if styles_root is None:
            return
        va = styles_root.find(f'./{W}docDefaults/{W}rPrDefault/{W}rPr/{W}vertAlign')
        if va is not None:
            self.default_vertical_alignment = va.get(W + 'val')
        rdef = styles_root.find(f"./{W}docDefaults/{W}rPrDefault/{W}rPr/{W}rFonts")
        if rdef is not None:
            self._apply_rfonts(rdef, self.default_fonts)
        for st in styles_root.findall(f"./{W}style"):
            sid = st.get(f"{W}styleId")
            if sid:
                self.styles[sid] = st
                if st.get(W + 'type') == 'paragraph' and st.get(W + 'default') in ('1', 'true', 'on'):
                    self.default_paragraph_style = sid

    def _resolve_theme_font(self, theme_name: str | None) -> str | None:
        if not theme_name:
            return None
        if theme_name in self.theme_fonts:
            return self.theme_fonts[theme_name]
        if theme_name.endswith('HAnsi') or theme_name.endswith('Ascii'):
            role = 'major' if theme_name.startswith('major') else 'minor'
            return self.theme_fonts.get(f'{role}HAnsi')
        if 'EastAsia' in theme_name:
            role = 'major' if theme_name.startswith('major') else 'minor'
            return self.theme_fonts.get(f'{role}EastAsia')
        return None

    def _apply_rfonts(self, rf: ET.Element, target: dict[str, str]) -> None:
        for attr, theme_attrs in (
            ('ascii', ('asciiTheme',)),
            ('hAnsi', ('hAnsiTheme',)),
            ('eastAsia', ('eastAsiaTheme',)),
            ('cs', ('cstheme', 'csTheme')),
        ):
            explicit = rf.get(f"{W}{attr}")
            if explicit:
                target[attr] = explicit
            else:
                for ta in theme_attrs:
                    tv = self._resolve_theme_font(rf.get(f"{W}{ta}"))
                    if tv:
                        target[attr] = tv
                        break

    def _chain(self, sid: str | None) -> list[ET.Element]:
        out: list[ET.Element] = []
        seen: set[str] = set()
        cur = sid
        while cur and cur not in seen and cur in self.styles:
            seen.add(cur)
            st = self.styles[cur]
            out.append(st)
            b = st.find(f"./{W}basedOn")
            cur = b.get(f"{W}val") if b is not None else None
        out.reverse()
        return out

    def _r_chain(self, r: ET.Element) -> list[ET.Element]:
        rs = r.find(f"./{W}rPr/{W}rStyle")
        return self._chain(rs.get(f"{W}val")) if rs is not None else []

    def effective_jc(self, p: ET.Element) -> str:
        jc = self.default_p_jc
        for st in self._chain(_p_style(p) or self.default_paragraph_style):
            el = st.find(f"./{W}pPr/{W}jc")
            if el is not None and el.get(f"{W}val"):
                jc = el.get(f"{W}val", jc)
        el = p.find(f"./{W}pPr/{W}jc")
        if el is not None and el.get(f"{W}val"):
            jc = el.get(f"{W}val", jc)
        return jc

    def effective_spacing(self, p: ET.Element) -> dict[str, str]:
        sp: dict[str, str] = {}
        for st in self._chain(_p_style(p) or self.default_paragraph_style):
            el = st.find(f"./{W}pPr/{W}spacing")
            if el is not None:
                for k in ("line", "lineRule", "before", "after"):
                    v = el.get(f"{W}{k}")
                    if v is not None:
                        sp[k] = v
        el = p.find(f"./{W}pPr/{W}spacing")
        if el is not None:
            for k in ("line", "lineRule", "before", "after"):
                v = el.get(f"{W}{k}")
                if v is not None:
                    sp[k] = v
        return sp

    def effective_ind(self, p: ET.Element) -> dict[str, str]:
        ind: dict[str, str] = {}
        for st in self._chain(_p_style(p) or self.default_paragraph_style):
            el = st.find(f"./{W}pPr/{W}ind")
            if el is not None:
                for k in ("firstLine", "firstLineChars", "left", "hanging"):
                    v = el.get(f"{W}{k}")
                    if v is not None:
                        ind[k] = v
        el = p.find(f"./{W}pPr/{W}ind")
        if el is not None:
            for k in ("firstLine", "firstLineChars", "left", "hanging"):
                v = el.get(f"{W}{k}")
                if v is not None:
                    ind[k] = v
        return ind

    def effective_bold(self, r: ET.Element, p: ET.Element | None = None) -> bool:
        val = False
        nodes = []
        if p is not None:
            nodes.extend(self._chain(_p_style(p) or self.default_paragraph_style))
        nodes.extend(self._r_chain(r))
        nodes.append(r)
        for node in nodes:
            rpr = node.find(f"./{W}rPr")
            if rpr is not None and rpr.find(f"./{W}b") is not None:
                b = rpr.find(f"./{W}b")
                val = b.get(f"{W}val", "1") not in ("0", "false", "off")
        return val

    def effective_italic(self, r: ET.Element, p: ET.Element | None = None) -> bool:
        val = False
        nodes = []
        if p is not None:
            nodes.extend(self._chain(_p_style(p) or self.default_paragraph_style))
        nodes.extend(self._r_chain(r))
        nodes.append(r)
        for node in nodes:
            rpr = node.find(f"./{W}rPr")
            if rpr is not None and rpr.find(f"./{W}i") is not None:
                i = rpr.find(f"./{W}i")
                val = i.get(f"{W}val", "1") not in ("0", "false", "off")
        return val

    def effective_ascii_font(self, r: ET.Element, p: ET.Element | None = None) -> str:
        resolved = dict(self.default_fonts)
        if p is not None:
            for st in self._chain(_p_style(p) or self.default_paragraph_style):
                rf = st.find(f"./{W}rPr/{W}rFonts")
                if rf is not None:
                    self._apply_rfonts(rf, resolved)
        for st in self._r_chain(r):
            rf = st.find(f"./{W}rPr/{W}rFonts")
            if rf is not None:
                self._apply_rfonts(rf, resolved)
        rf = r.find(f"./{W}rPr/{W}rFonts")
        if rf is not None:
            self._apply_rfonts(rf, resolved)
        return resolved.get("ascii") or resolved.get("hAnsi") or ""

    def effective_eastasia_font(self, r: ET.Element, p: ET.Element | None = None) -> str | None:
        resolved = dict(self.default_fonts)
        if p is not None:
            for st in self._chain(_p_style(p) or self.default_paragraph_style):
                rf = st.find(f"./{W}rPr/{W}rFonts")
                if rf is not None:
                    self._apply_rfonts(rf, resolved)
        for st in self._r_chain(r):
            rf = st.find(f"./{W}rPr/{W}rFonts")
            if rf is not None:
                self._apply_rfonts(rf, resolved)
        rf = r.find(f"./{W}rPr/{W}rFonts")
        if rf is not None:
            self._apply_rfonts(rf, resolved)
        return resolved.get("eastAsia")


def _para_num_info(p: ET.Element, styles: StyleResolver, numbering: tuple[dict[str, str], dict[tuple[str, str], tuple[str, str]]]) -> tuple[str, str, str] | None:
    num_to_abs, levels = numbering
    num_id = ilvl = None
    for st in styles._chain(_p_style(p) or styles.default_paragraph_style):
        np = st.find(f"./{W}pPr/{W}numPr")
        if np is not None:
            ni = np.find(W + 'numId')
            il = np.find(W + 'ilvl')
            if ni is not None and ni.get(W + 'val'):
                num_id = ni.get(W + 'val')
            if il is not None and il.get(W + 'val'):
                ilvl = il.get(W + 'val')
    np = p.find(f"./{W}pPr/{W}numPr")
    if np is not None:
        ni = np.find(W + 'numId')
        il = np.find(W + 'ilvl')
        if ni is not None and ni.get(W + 'val'):
            num_id = ni.get(W + 'val')
        if il is not None and il.get(W + 'val'):
            ilvl = il.get(W + 'val')
    if not num_id or num_id == '0':
        return None
    abs_id = num_to_abs.get(num_id)
    if not abs_id:
        return None
    lvl_text, suff = levels.get((abs_id, ilvl or '0'), ('', 'tab'))
    return num_id, lvl_text, suff


def _is_front_title(txt: str) -> bool:
    return bool(re.match(r"^(?:摘\s*要|中文摘要|abstract|目\s*录|contents|版权.*声明|致\s*谢|个人简历|附\s*录|参考\s*文献)\b", txt.strip(), re.I))


def _is_super_run(r: ET.Element) -> bool:
    va = r.find(f"./{W}rPr/{W}vertAlign")
    return va is not None and va.get(f"{W}val") == "superscript"


def _is_sub_run(r: ET.Element) -> bool:
    va = r.find(f"./{W}rPr/{W}vertAlign")
    return va is not None and va.get(f"{W}val") == "subscript"


def _expand_num_spec(spec: str) -> list[int]:
    nums: list[int] = []
    norm = spec.strip().strip("[]").replace("，", ",").replace("~", "-").replace("–", "-").replace("—", "-")
    for part in re.split(r"\s*,\s*", norm):
        if not part:
            continue
        if "-" in part:
            m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)
            if m:
                a, b = int(m.group(1)), int(m.group(2))
                if a <= b and (b - a) < 100:
                    nums.extend(range(a, b + 1))
        elif part.isdigit():
            nums.append(int(part))
    return nums


def _is_non_citation_bracket(raw_spec: str) -> bool:
    norm = raw_spec.strip().replace("，", ",")
    if "," in norm:
        parts = [p.strip() for p in norm.split(",") if p.strip()]
        if parts == ["0", "1"]:
            return True
    if re.fullmatch(r"(?:19|20)\d{2}", norm):
        return True
    return False


def _detect_mode(all_text: str, explicit: str) -> str:
    if explicit != "auto":
        return explicit
    head = all_text[:1500]
    if "开题报告" in head:
        return "proposal"
    if "中期考核" in head:
        return "midterm"
    if "轮转报告" in head:
        return "rotation"
    if "硕士研究生学位论文" in head:
        return "master-thesis"
    return "phd-thesis"


def _extract_runs_skip_toc_fields(p: ET.Element, field_depth: int = 0) -> tuple[list[ET.Element], int]:
    out: list[ET.Element] = []
    depth = field_depth
    for child in list(p):
        tag = _tag_local(child)
        if tag == "hyperlink" and depth == 0:
            out.extend(list(child.findall(f"./{W}r")))
        elif tag == "r":
            fc = child.find(f"./{W}fldChar")
            if fc is not None:
                ft = fc.get(f"{W}fldCharType")
                if ft == "begin":
                    depth += 1
                elif ft == "end" and depth > 0:
                    depth -= 1
            elif depth == 0:
                out.append(child)
    return out, depth


def _caption_runs(p: ET.Element) -> list[ET.Element]:
    runs: list[ET.Element] = []
    for child in list(p):
        tag = _tag_local(child)
        if tag == "r" and _r_text(child):
            runs.append(child)
        elif tag == "fldSimple":
            for r in child.findall(f"./{W}r"):
                if _r_text(r):
                    runs.append(r)
    return runs


def run_audit(args: argparse.Namespace) -> tuple[dict, list[CheckResult]]:
    if not str(args.docx or "").strip():
        raise ValueError("--docx path cannot be empty")
    docx_path = pathlib.Path(args.docx).resolve()
    if not docx_path.is_file():
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    preset = load_effective_preset(args.preset, args.profile_config, None, None, args.citation_style)
    citation_style = preset.get("citation_style", "nature")
    if citation_style not in ("nature", "gb7714"):
        raise ValueError(f"Unsupported citation_style {citation_style!r}; expected 'nature' or 'gb7714'")
    template = _template_snapshot(args.template) if args.template else None
    typography = preset.get('typography', {})

    results: list[CheckResult] = []
    stats: dict[str, object] = {
        "docx": str(docx_path),
        "sha256": _sha256(docx_path.read_bytes()),
        "preset": preset["preset_name"],
        "citation_style": citation_style,
        "template": str(pathlib.Path(args.template).resolve()) if template else None,
    }

    with zipfile.ZipFile(docx_path, "r") as zf:
        bad_zip = zf.testzip()
        results.append(
            CheckResult(
                "Pkg-CRC",
                "工程基础-包完整性",
                "PASS" if bad_zip is None else "FAIL",
                "ZIP CRC 校验通过" if bad_zip is None else f"ZIP 损坏成员: {bad_zip}",
                [] if bad_zip is None else [str(bad_zip)],
            )
        )

        names = set(zf.namelist())
        xml_errors: list[str] = []
        parsed: dict[str, ET.Element] = {}
        for name in sorted(names):
            if name.endswith((".xml", ".rels")):
                try:
                    parsed[name] = ET.fromstring(zf.read(name))
                except Exception as exc:
                    xml_errors.append(f"{name}: {exc}")
        if len(names) != len(zf.namelist()):
            xml_errors.append('ZIP 中存在重复部件路径')
        try:
            word_main_type(parsed.get('[Content_Types].xml'))
        except ValueError as exc:
            xml_errors.append(str(exc))
        main = parsed.get('word/document.xml')
        if main is None or main.tag != W + 'document' or main.find(W + 'body') is None:
            xml_errors.append('缺少有效的 Word 主文档及正文')
        results.append(
            CheckResult(
                "Pkg-XML",
                "工程基础-XML语法",
                "PASS" if not xml_errors else "FAIL",
                f"共解析 {len(parsed)} 个 XML/RELS 部件，无语法错误" if not xml_errors else f"XML 解析失败 {len(xml_errors)} 处",
                xml_errors[:10],
            )
        )
        if xml_errors or "word/document.xml" not in parsed:
            return stats, results

        # Relationship closure
        broken_rels: list[str] = []
        if not has_main_relationship(parsed.get('_rels/.rels')):
            broken_rels.append('_rels/.rels must contain one internal officeDocument relationship to word/document.xml')
        rels_root = parsed.get("word/_rels/document.xml.rels")
        rels_map: dict[str, str] = {}
        if rels_root is not None:
            for rel in list(rels_root):
                rid = rel.get("Id", "")
                tgt = rel.get("Target", "")
                if rel.get("TargetMode") == "External":
                    continue
                norm = resolve_target('word/document.xml', tgt)
                rels_map[rid] = norm
                if norm not in names:
                    broken_rels.append(f"{rid} -> {tgt} ({norm})")
        results.append(
            CheckResult(
                "Pkg-Rels",
                "工程基础-关系闭环",
                "PASS" if not broken_rels else "FAIL",
                "Word 主文档入口和主文档内部关系目标通过检查" if not broken_rels else f"发现 {len(broken_rels)} 项 Word 包关系问题",
                broken_rels[:10],
            )
        )

        doc_root = parsed["word/document.xml"]
        anchors = {W + tag for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference')}
        changes = {W + tag for tag in ('ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange',
                                        'sectPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange',
                                        'cellIns', 'cellDel', 'cellMerge', 'numberingChange')}
        word_tags = [el.tag for name, root in parsed.items() if name.startswith('word/') for el in root.iter()]
        comment_anchors = sum(tag in anchors for tag in word_tags)
        comment_items = word_tags.count(W + 'comment')
        revisions = sum(tag in changes for tag in word_tags)
        results.append(
            CheckResult(
                "Pkg-Clean",
                "工程基础-批注与修订清理",
                "PASS" if (not comment_anchors and not comment_items and not revisions) else "FAIL",
                f"批注标记={comment_anchors}, 批注条数={comment_items}, 修订标记={revisions}",
                [] if (not comment_anchors and not comment_items and not revisions) else ["文档残留未接受修订或未删除批注"],
            )
        )

        if args.base_docx:
            base_path = pathlib.Path(args.base_docx).resolve()
            with zipfile.ZipFile(base_path, "r") as zbase:
                base_media = {n: _sha256(zbase.read(n)) for n in zbase.namelist() if n.startswith("word/media/")}
            cand_media = {n: _sha256(zf.read(n)) for n in names if n.startswith("word/media/")}
            media_ok = base_media == cand_media
            results.append(
                CheckResult(
                    "Pkg-Media",
                    "工程基础-媒体二进制零损",
                    "PASS" if media_ok else "FAIL",
                    f"word/media 数量={len(cand_media)}，哈希与基线完全一致" if media_ok else "word/media 与基线存在新增或变更",
                    [],
                )
            )

        styles = StyleResolver(parsed.get("word/styles.xml"), parsed.get("word/theme/theme1.xml"))
        numbering_defs = _load_numbering(parsed.get("word/numbering.xml"))
        body = doc_root.find(f"./{W}body")
        assert body is not None
        results.extend(_direct_format_checks(preset.get('audit', {}), styles, body))
        blocks = list(_iter_story_blocks(body))
        all_paras = list(body.iter(f"{W}p"))
        body_paras = [el for el in blocks if el.tag == f"{W}p"]
        all_text = "\n".join(_p_text(p) for p in all_paras)

        mode = _detect_mode(all_text, args.mode)
        stats["detected_mode"] = mode

        # 1. Settings & page margins (from preset)
        settings_root = parsed.get("word/settings.xml")
        page_cfg = preset.get("page", {})
        flags = ('mirrorMargins', 'evenAndOddHeaders')
        settings = lambda root: [_on(root.find(W + flag)) if root is not None else False for flag in flags]
        want_settings = settings(template['settings']) if template else [
            bool(page_cfg.get('mirror_margins', True)), bool(page_cfg.get('odd_even_headers', True))]
        actual_settings = settings(settings_root)
        results.append(
            CheckResult(
                "Layout-Settings",
                "排版-1(对称页边距/奇偶页)",
                "PASS" if actual_settings == want_settings else "FAIL",
                f"mirrorMargins={actual_settings[0]}, evenAndOddHeaders={actual_settings[1]} (期望 {want_settings})",
                [],
            )
        )

        exp_dims = _resolve_page_dxa(page_cfg)
        exp_dims['w'] = exp_dims.pop('width_dxa')
        exp_dims['h'] = exp_dims.pop('height_dxa')
        exp_dims['orient'] = 'portrait' if exp_dims['w'] <= exp_dims['h'] else 'landscape'
        sect_prs = _live_sections(doc_root)
        draft_geom = [_section_geometry(sp) for sp in sect_prs]
        numbering = _section_numbering(sect_prs, parsed, rels_map)
        draft_sec_meta = _section_meta(draft_geom, numbering)
        margin_issues: list[str] = []
        sec_mappings: list[dict] = []

        if template and sect_prs:
            sec_mappings = align_section_indices(draft_sec_meta, template['section_meta'])
            if len(sect_prs) != len(template['geometry']):
                used_t_indices = {m['template_idx'] for m in sec_mappings}
                required_roles = {
                    i for i, tm in enumerate(template['section_meta'])
                    if tm['fmt'] in ('upperRoman', 'lowerRoman') or (tm['fmt'] == 'decimal' and tm['has_page'])
                }
                if len(sect_prs) < len(template['geometry']) or not required_roles.issubset(used_t_indices):
                    margin_issues.append(f"分节数 {len(sect_prs)} 与模板 {len(template['geometry'])} 不兼容")
            for item in sec_mappings:
                s_idx = item['draft_idx'] + 1
                ti = item['template_idx']
                actual_geom = draft_geom[item['draft_idx']]
                dimensions = dict(template['geometry'][ti])
                if item['preserve_landscape']:
                    tw, th = dimensions.get('w'), dimensions.get('h')
                    if isinstance(tw, int) and isinstance(th, int):
                        dimensions['w'] = max(tw, th)
                        dimensions['h'] = min(tw, th)
                    dimensions['orient'] = 'landscape'
                for attr, actual in actual_geom.items():
                    want = dimensions[attr]
                    if attr == 'gutter' and (actual or 0) == (want or 0):
                        continue
                    if actual == want:
                        continue
                    unit_note = '' if attr == 'orient' else f" ({_fmt_dxa(actual)} vs 期望 {_fmt_dxa(want)})"
                    margin_issues.append(f"第 {s_idx} 节: {attr}={actual} (期望 {want}){unit_note}")
        elif sect_prs:
            for s_idx, actual_geom in enumerate(draft_geom, 1):
                target_dims = dict(exp_dims)
                if s_idx > 1 and actual_geom.get('orient') == 'landscape' and (
                    actual_geom.get('w'), actual_geom.get('h')
                ) == (exp_dims['h'], exp_dims['w']):
                    target_dims['w'], target_dims['h'] = exp_dims['h'], exp_dims['w']
                    target_dims['orient'] = 'landscape'
                for attr, actual in actual_geom.items():
                    want = target_dims[attr]
                    if attr == 'gutter' and (actual or 0) == (want or 0):
                        continue
                    if actual == want:
                        continue
                    if not isinstance(actual, int) or not isinstance(want, int) or abs(actual - want) > 25:
                        margin_issues.append(f"第 {s_idx} 节: {attr}={_fmt_dxa(actual)} (期望 {_fmt_dxa(want)})")
        if not sect_prs:
            margin_issues.append('缺少页面设置分节')

        results.append(
            CheckResult(
                "Layout-1",
                "排版-1(纸张与页边距)",
                "PASS" if (sect_prs and not margin_issues) else "FAIL",
                (
                    f"全部 {len(sect_prs)} 节页面设置符合{'模板' if template else '当前预设'}"
                    if not margin_issues
                    else f"发现 {len(margin_issues)} 处页面偏差"
                ),
                margin_issues[:12],
            )
        )

        # 3-section & header isolation
        sec_issues: list[str] = []
        sec_descs = [f'S{i}({fmt}@{start or "续编"}, PAGE={",".join(active) or "-"})'
                     for i, (fmt, start, _, _, active, _) in enumerate(numbering, 1)]
        if template:
            if _direct_equal_length_compatible(draft_sec_meta, template['section_meta']):
                for s_idx, (got, exp) in enumerate(zip(numbering, template['numbering']), 1):
                    if got != exp:
                        for label, g_val, e_val in (
                            ('页码格式', got[0], exp[0]),
                            ('起始页码', got[1] or '续编', exp[1] or '续编'),
                            ('分节起始方式', got[2], exp[2]),
                            ('首页不同', got[3], exp[3]),
                            ('页脚PAGE域', got[4], exp[4]),
                            ('页眉类型', got[5], exp[5]),
                        ):
                            if g_val != e_val:
                                sec_issues.append(f"第 {s_idx} 节 {label}: 实际 {g_val} ≠ 模板 {e_val}")
            elif sec_mappings:
                for item in sec_mappings:
                    di, ti = item['draft_idx'], item['template_idx']
                    got, exp = numbering[di], template['numbering'][ti]
                    s_idx = di + 1
                    if got[0] != exp[0]:
                        sec_issues.append(f"第 {s_idx} 节 页码格式: 实际 {got[0]} ≠ 模板 {exp[0]}")
                    exp_start = (
                        exp[1]
                        if (item['role'] == 'roman_front' or (item['role'] == 'decimal_body' and item['is_first_body']))
                        else None
                    )
                    if item['role'] == 'landscape_island' and got[1] is not None:
                        sec_issues.append(f"第 {s_idx} 节（横向宽表节）不应重新起编页码 (start={got[1]})")
                    elif item['role'] != 'landscape_island' and got[1] != exp_start:
                        sec_issues.append(f"第 {s_idx} 节 起始页码: 实际 {got[1] or '续编'} ≠ 模板 {exp_start or '续编'}")
                    if got[2] != exp[2]:
                        sec_issues.append(f"第 {s_idx} 节 分节起始方式: 实际 {got[2]} ≠ 模板 {exp[2]}")
                    if item['role'] != 'landscape_island' and got[3] != exp[3]:
                        sec_issues.append(f"第 {s_idx} 节 首页不同: 实际 {got[3]} ≠ 模板 {exp[3]}")
                    if bool(got[4]) != bool(exp[4]):
                        sec_issues.append(f"第 {s_idx} 节 页脚PAGE域: 实际 {got[4]} ≠ 模板 {exp[4]}")
                    if bool(got[5]) != bool(exp[5]):
                        sec_issues.append(f"第 {s_idx} 节 页眉类型: 实际 {got[5]} ≠ 模板 {exp[5]}")
            else:
                sec_issues.append('分节起始方式、首页设置、页眉类型、页码格式或页码域与模板不同')
            actual_styles = _style_properties(parsed.get('word/styles.xml'))
            changed = [key for key, props in template['style_props'].items() if actual_styles.get(key) != props]
            results.append(CheckResult('Layout-TemplateStyles', '模板样式定义', 'WARN' if changed else 'PASS',
                                       f'模板样式属性有 {len(changed)} 项差异' if changed else '模板样式属性匹配', changed[:15]))
        elif len(sect_prs) < 3:
            sec_issues.append(f"仅发现 {len(sect_prs)} 节；规范应至少分 3 节（封面无页码 / 摘要目录罗马 / 正文阿拉伯）")
        else:
            roman_idx = next((i for i, row in enumerate(numbering) if row[0] in ('upperRoman', 'lowerRoman')), None)
            dec_idx = next(
                (i for i, row in enumerate(numbering)
                 if row[0] == 'decimal' and row[1] == '1' and row[4] and (roman_idx is None or i > roman_idx)),
                None,
            )
            front_end = roman_idx if roman_idx is not None else 1
            for fi in range(front_end):
                for hr in sect_prs[fi].findall(f"./{W}headerReference"):
                    tgt = rels_map.get(hr.get(f"{R}id", ""), "")
                    if tgt in parsed and "".join(t.text or "" for t in parsed[tgt].iter(f"{W}t")).strip():
                        sec_issues.append(f"第 {fi + 1} 节（封面/声明）不应显示页眉文字")
                if numbering[fi][4]:
                    sec_issues.append(f"第 {fi + 1} 节（封面/声明）的页脚包含 PAGE 页码域")
            if roman_idx is None or roman_idx == 0 or numbering[roman_idx][:2] != ('upperRoman', '1') or not numbering[roman_idx][4]:
                sec_issues.append('摘要/目录节应从大写罗马 I 起编，并使用 PAGE 页码域')
            if dec_idx is None or (roman_idx is not None and dec_idx <= roman_idx):
                sec_issues.append("正文起始/主体节缺少显式 pgNumType/@w:start='1' 的阿拉伯数字页码")
            else:
                for tail_i in range(dec_idx + 1, len(numbering)):
                    if numbering[tail_i][1] == '1':
                        sec_issues.append(f"第 {tail_i + 1} 节正文续节不应重复设置 start='1' 导致页码重置")
        results.append(
            CheckResult(
                "Layout-SecNum",
                "排版-2(三节页码与页眉隔离)",
                "PASS" if not sec_issues else "FAIL",
                f"共 {len(sect_prs)} 节：" + " -> ".join(sec_descs),
                sec_issues,
            )
        )

        # 2. TOC
        has_toc = _instructions(doc_root, 'TOC')
        toc1_st = styles.styles.get("TOC1")
        toc1_bold = toc1_st is not None and _on(toc1_st.find(f"./{W}rPr/{W}b"))
        toc_placeholder = any("更新域" in _p_text(p) for p in all_paras[:150])
        toc_status = "PASS"
        toc_details: list[str] = []
        if not has_toc and (not template or template['toc']):
            toc_status = "FAIL"
            toc_details.append("未找到 Word TOC 目录域代码")
        elif toc_placeholder:
            toc_status = "WARN"
            toc_details.append("目录仍为占位文本，交付前须在 Word 中右键更新目录域 (F9)")
        if not toc1_bold and not template:
            toc_status = "FAIL"
            toc_details.append("styles.xml 中 TOC1 一级目录样式未设加粗")
        if template and not has_toc and not template['toc']:
            toc_status = 'INFO'
        results.append(
            CheckResult(
                "Layout-TOC",
                "排版&常见错误-1(目录域与TOC1加粗)",
                toc_status,
                {'PASS': '已检测目录域', 'WARN': '目录域已配置（待 Word F9 更新）',
                 'FAIL': '目录配置存在异常', 'INFO': '模板未要求目录域'}[toc_status],
                toc_details,
            )
        )

        # 3. Heading spacing & paragraph alignment & Latin font checks
        heading_space_issues: list[str] = []
        align_issues: list[str] = []
        font_issues: list[str] = []
        spacing_issues: list[str] = []
        body_spacing = str(round(20 * float(typography.get('body_line_pt', 20))))
        latin_fonts = {typography.get('latin_font', 'Times New Roman').casefold(), 'arial'}
        body_checked = 0
        field_depth = 0
        for p in body_paras:
            ps = _p_style(p)
            txt = _p_text(p).strip()
            runs_outside_fields, field_depth = _extract_runs_skip_toc_fields(p, field_depth)
            if not txt:
                continue
            num_info = _para_num_info(p, styles, numbering_defs)
            if ps in (CH1_STYLES | H2_STYLES | H3_STYLES | H4_STYLES):
                m_h = re.match(r"^(第[一二三四五六七八九十\d]+章|\d+(?:\.\d+){1,3})([\s\u3000]*)(.+)$", txt)
                if m_h and m_h.group(2) != " ":
                    heading_space_issues.append(f"标题序号空格非单半角: {txt[:35]!r}")
                elif m_h is None and num_info is not None and num_info[2] == 'nothing' and not _p_text(p).startswith(" "):
                    heading_space_issues.append(f"自动编号标题后缺少间距 ({num_info[1]}): {txt[:35]!r}")
            if ps in (BODY_STYLES | ABS_ZH_STYLES | ABS_EN_STYLES | FIG_CAP_STYLES | BIB_STYLES):
                jc = styles.effective_jc(p)
                if jc != "both" and "签字" not in txt and "签名" not in txt and "声明" not in txt and not txt.startswith("在学期间"):
                    align_issues.append(f"样式 {ps} 非两端对齐 ({jc}): {txt[:30]!r}")
            if ps in ('BodyTextZH', '08'):
                body_checked += 1
                space = styles.effective_spacing(p)
                if (space.get('line'), space.get('lineRule')) != (body_spacing, 'exact'):
                    spacing_issues.append(f'{ps}: {space.get("line")}/{space.get("lineRule")}，期望 {body_spacing}/exact: {txt[:30]}')
            if ps in (CH1_STYLES | H2_STYLES | H3_STYLES | H4_STYLES | BODY_STYLES | ABS_ZH_STYLES | ABS_EN_STYLES | FIG_CAP_STYLES | TBL_TITLE_STYLES | TBL_NOTE_STYLES | BIB_STYLES):
                for r in runs_outside_fields:
                    rt = _r_text(r)
                    if not rt or rt.isspace() or "【" in rt:
                        continue
                    if re.search(r"[A-Za-z0-9α-ωΑ-Ωμµ°]", rt):
                        af_raw = styles.effective_ascii_font(r, p)
                        af = af_raw.lower()
                        if not af or af not in latin_fonts:
                            font_issues.append(f"英文字体未设或不在当前允许字体中 ({af_raw or '未设'}): {rt[:30]!r} @ {ps}")

        results.append(CheckResult("Err-9", "常见错误-9(标题序号与标题间单半角空格)", "PASS" if not heading_space_issues else "FAIL", "各级标题序号后空格符合单半角空格要求" if not heading_space_issues else f"发现 {len(heading_space_issues)} 处标题空格异常", heading_space_issues[:10]))
        results.append(CheckResult("Layout-9", "排版-9(正文/摘要/图注/参考文献两端对齐)", "PASS" if not align_issues else "FAIL", "全部核查段落均设为两端对齐" if not align_issues else f"发现 {len(align_issues)} 处未两端对齐", align_issues[:10]))
        results.append(CheckResult("Err-8", "常见错误-8(英文数字与符号用英文字体)", "PASS" if not font_issues else "FAIL", "未发现英文/数字/希腊字母/°C 落入中文字体" if not font_issues else f"发现 {len(font_issues)} 处英文字体偏差", font_issues[:12]))
        results.append(CheckResult('Layout-BodySpacing', '正文固定行距',
                                   'FAIL' if spacing_issues else 'PASS' if body_checked else 'INFO',
                                   f'正文命中已知样式 {body_checked} 段，固定行距偏差 {len(spacing_issues)} 段', spacing_issues[:12]))

        # Locate chapter 1 -> references -> back matter boundaries (literal or Word auto-numbering; exclude TOC entries)
        ch1_idx = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and not _is_toc_para(el) and re.match(r"^(?:第一章|Chapter\s+1)\b", _p_text(el).strip(), re.I)), None)
        if ch1_idx is None:
            for i, el in enumerate(blocks):
                if el.tag != f"{W}p" or _is_toc_para(el):
                    continue
                ps = _p_style(el)
                txt = _p_text(el).strip()
                if not txt or _is_front_title(txt):
                    continue
                ninfo = _para_num_info(el, styles, numbering_defs)
                if (ps in CH1_STYLES and ninfo is not None) or (ninfo is not None and ("第" in ninfo[1] and "章" in ninfo[1] or "Chapter" in ninfo[1])):
                    ch1_idx = i
                    break
        ref_idx = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and not _is_toc_para(el) and re.fullmatch(r"参考\s*文献|References|Bibliography|Works Cited", _p_text(el).strip(), re.I)), None)
        cv_idx = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and not _is_toc_para(el) and _p_text(el).strip().replace("\u3000", "").replace(" ", "").startswith(("个人简历", "作者简历", "作者简介", "附录", "致谢"))), None)
        body_start = ch1_idx or 0
        body_end = ref_idx or len(blocks)

        main_blocks = blocks[body_start:body_end]
        main_paras = [p for el in main_blocks for p in ([el] if el.tag == f"{W}p" else list(el.iter(f"{W}p")))]

        # 4. Figures, Tables, Three-Line Borders & Closure
        fig_captions: list[tuple[str, int, ET.Element]] = []
        tbl_captions: list[tuple[str, int, ET.Element]] = []
        caption_bold_issues: list[str] = []
        bare_caption_issues: list[str] = []
        lower_panel_issues: list[str] = []
        adapted_fig_issues: list[str] = []

        for p in main_paras:
            ps = _p_style(p)
            txt = _p_text(p).strip()
            if not txt:
                continue
            mf = re.match(r"^(?:图|补充图)\s*([1-9]\d*|[A-D])[.\-](\d+)\s+(.+)$", txt)
            if mf and (ps in FIG_CAP_STYLES or re.match(r"^(?:图|补充图)\s*(?:[1-9]\d*|[A-D])[.\-]\d+\s+\S+", txt)) and len(txt) < 1200:
                ch, num = mf.group(1), int(mf.group(2))
                fig_captions.append((ch, num, p))
                runs = _caption_runs(p)
                if runs and not styles.effective_bold(runs[0], p):
                    caption_bold_issues.append(f"图 {ch}.{num} 题名首段未加粗")
                if len(runs) > 1 and styles.effective_bold(runs[-1], p) and len(txt) > 35:
                    caption_bold_issues.append(f"图 {ch}.{num} 详细说明误全文加粗")
                if re.search(r"(?:^|[。.\s;])([a-z])(?:\)|\.|\s*[：:])", mf.group(3)):
                    lower_panel_issues.append(f"图 {ch}.{num} 出现小写分图号")
                if len(mf.group(3)) < 18 and "修改自" not in txt:
                    bare_caption_issues.append(f"图 {ch}.{num} 图注过于简略")
                if "修改自" in txt and not re.search(r"修改自[^。]+。?$", txt):
                    adapted_fig_issues.append(f"图 {ch}.{num} 改编图“修改自…”须置于图注最末尾")

            mt = re.match(r"^(?:表|附表)\s*([1-9]\d*|[A-D])[.\-](\d+)\s+(.+)$", txt)
            is_tbl_title = (ps in TBL_TITLE_STYLES) or (mt is not None and not re.search(r"[。；，]\s*$", txt) and "。" not in txt[:35] and len(txt) < 120)
            if mt and is_tbl_title:
                ch, num = mt.group(1), int(mt.group(2))
                tbl_captions.append((ch, num, p))
                runs = _caption_runs(p)
                if runs and not styles.effective_bold(runs[0], p):
                    caption_bold_issues.append(f"表 {ch}.{num} 题名未加粗")

        def _seq_check(items: list[tuple[str, int, ET.Element]], kind: str) -> list[str]:
            by_ch: dict[str, list[int]] = {}
            for ch, num, _ in items:
                by_ch.setdefault(ch, []).append(num)
            errs: list[str] = []
            for ch, nums in by_ch.items():
                if nums != list(range(1, len(nums) + 1)):
                    errs.append(f"{kind} {ch}.x 编号不连续: {nums}")
            return errs

        seq_issues = _seq_check(fig_captions, "图") + _seq_check(tbl_captions, "表")
        fig_cap_set = {(ch, n) for ch, n, _ in fig_captions}
        tbl_cap_set = {(ch, n) for ch, n, _ in tbl_captions}
        fig_cite_set: set[tuple[str, int]] = set()
        tbl_cite_set: set[tuple[str, int]] = set()
        cap_elem_set = {id(p) for _, _, p in fig_captions + tbl_captions}
        for p in main_paras:
            if id(p) in cap_elem_set:
                continue
            txt = _p_text(p)
            for m in re.finditer(r"(?:图|补充图)\s*([1-9]\d*|[A-D])[.\-](\d+)", txt):
                fig_cite_set.add((m.group(1), int(m.group(2))))
            for m in re.finditer(r"(?:表|附表)\s*([1-9]\d*|[A-D])[.\-](\d+)", txt):
                tbl_cite_set.add((m.group(1), int(m.group(2))))
        closure_issues: list[str] = []
        if fig_cap_set - fig_cite_set:
            closure_issues.append(f"正文未引用图: {sorted(fig_cap_set - fig_cite_set)}")
        if fig_cite_set - fig_cap_set:
            closure_issues.append(f"正文引用悬空图编号: {sorted(fig_cite_set - fig_cap_set)}")
        if tbl_cap_set - tbl_cite_set:
            closure_issues.append(f"正文未引用表: {sorted(tbl_cap_set - tbl_cite_set)}")
        if tbl_cite_set - tbl_cap_set:
            closure_issues.append(f"正文引用悬空表编号: {sorted(tbl_cite_set - tbl_cap_set)}")

        # Three-line tables inside main_blocks (skip placeholder dashed figure frames)
        bad_tables: list[str] = []
        unresolved_tables: list[str] = []
        content_tbl_count = 0
        for tidx, tbl in enumerate([el for el in main_blocks if el.tag == f"{W}tbl"], 1):
            bdr = tbl.find(f"./{W}tblPr/{W}tblBorders")
            top_b = bdr.find(f"./{W}top") if bdr is not None else None
            # The bundled starters use one dashed cell as a picture placeholder.
            # Keep real one-cell tables in scope when they contain other text.
            cells = tbl.findall(f"./{W}tr/{W}tc")
            placeholder_labels = ('此处插入矢量科研图', '在此放置作品、资料或研究示意图', '在此放置高分辨率矢量科研图')
            if (len(cells) == 1 and top_b is not None and top_b.get(f"{W}val") == 'dashed'
                    and _p_text(tbl).lstrip('【[ ').startswith(placeholder_labels)):
                continue
            content_tbl_count += 1
            issues, warnings, needs_word = _table_problems(tbl)
            if issues:
                bad_tables.extend(f'表 #{tidx}: {item}' for item in issues)
            if warnings or needs_word:
                unresolved_tables.extend(f'表 #{tidx}: {item}' for item in warnings)

        min_fig_pt = float(typography.get("figure_min_font_pt", 7.5))
        fig_font_checked, fig_font_issues = _inspect_figure_font_sizes(main_blocks, zf, names, rels_map, min_fig_pt)
        fig5_status = "FAIL" if fig_font_issues else ("PASS" if fig_font_checked else "INFO")
        fig5_summary = (
            f"检出 {len(fig_font_issues)} 处图内文字小于 6 号字 ({min_fig_pt:g} pt)"
            if fig_font_issues
            else (
                f"已核查 {fig_font_checked} 处图内矢量文字，均不小于 6 号字 ({min_fig_pt:g} pt)"
                if fig_font_checked
                else f"未检出可机读矢量图文字，请在成稿 PDF 复核图内文字不小于 6 号字 ({min_fig_pt:g} pt)"
            )
        )
        results.append(CheckResult("Fig-1", "图表要求-1(按章连续编号)", "PASS" if not seq_issues else "FAIL", f"图={len(fig_captions)}, 表={len(tbl_captions)}", seq_issues))
        results.append(CheckResult("Fig-3-8", "图表要求-3/8(正文图表双向闭环)", "PASS" if not closure_issues else "FAIL", "全部图表与正文双向引用闭环" if not closure_issues else "存在未引用或悬空图表编号", closure_issues))
        results.append(CheckResult("Fig-4", "图表要求-4(图注/表题样式与局部加粗)", "PASS" if not caption_bold_issues else "FAIL", "图注与表题局部加粗规范" if not caption_bold_issues else f"发现 {len(caption_bold_issues)} 处标题加粗问题", caption_bold_issues[:10]))
        results.append(CheckResult("Fig-5", "图表要求-5(图内文字不小于6号字)", fig5_status, fig5_summary, fig_font_issues[:10]))
        results.append(CheckResult("Fig-9-10", "图表要求-9/10(大写分图号与完整图注)", "PASS" if (not lower_panel_issues and not bare_caption_issues) else "FAIL", "分图号大写且图注说明完整" if (not lower_panel_issues and not bare_caption_issues) else "存在小写分图号或过短图注", (lower_panel_issues + bare_caption_issues)[:10]))
        results.append(CheckResult("Acad-2", "学术规范-2(改编图出处置图注末尾)", "PASS" if not adapted_fig_issues else "FAIL", "改编图规范检查通过" if not adapted_fig_issues else f"发现 {len(adapted_fig_issues)} 处改编图标注问题", adapted_fig_issues))
        table_status = 'FAIL' if bad_tables else 'WARN' if unresolved_tables else 'PASS' if content_tbl_count else 'INFO'
        results.append(CheckResult('Tbl-3Line', '三线表边框规范', table_status,
                                   f'识别 {content_tbl_count} 个内容表格，边框异常 {len(bad_tables)} 处，待确认 {len(unresolved_tables)} 处',
                                   (bad_tables + unresolved_tables)[:15]))

        # 5. Citations & Bibliography (Nature vs GB/T 7714)
        cited_nums: set[int] = set()
        cite_groups = 0
        bad_cite_place: list[str] = []
        prose_end = min((i for i in (cv_idx, ref_idx) if i is not None and i >= body_start), default=len(blocks))
        prose_blocks = blocks[body_start:prose_end]
        prose_paras = [p for el in prose_blocks for p in ([el] if el.tag == f"{W}p" else list(el.iter(f"{W}p")))]

        for p in prose_paras:
            ps = _p_style(p)
            if ps in BIB_STYLES:
                continue
            txt = _p_text(p)
            for child in list(p):
                if _tag_local(child) == "hyperlink" and (child.get(f"{W}anchor") or "").startswith("_BIB_"):
                    m_b = re.match(r"_BIB_(\d+)", child.get(f"{W}anchor", ""))
                    if m_b:
                        cited_nums.add(int(m_b.group(1)))
            if citation_style == "nature":
                runs = p.findall(f"./{W}r")
                in_cwyw = False
                prev_visible = ""
                for ri, r in enumerate(runs):
                    fc = r.find(f"./{W}fldChar")
                    instr = r.find(f"./{W}instrText")
                    if instr is not None and "EN.CITE" in (instr.text or "").upper():
                        in_cwyw = True
                    if fc is not None:
                        ft = fc.get(f"{W}fldCharType")
                        if ft == "end":
                            in_cwyw = False
                        continue
                    rt_raw = _r_text(r) or ""
                    if _is_super_run(r) and not in_cwyw:
                        rt = rt_raw.strip()
                        nums = _expand_num_spec(rt)
                        if nums and not re.match(r"^[\d,，\-–—~\s]+$", rt) and not rt.startswith("["):
                            nums = []
                        if nums:
                            cite_groups += 1
                            cited_nums.update(nums)
                            if prev_visible and prev_visible[-1] in "。，；！？.,;!?" and ps not in TBL_NOTE_STYLES:
                                bad_cite_place.append(f"标点后出现引用上标 {rt!r}: {prev_visible[-15:]!r}")
                    if rt_raw:
                        prev_visible += rt_raw
                for m_br in re.finditer(r"\[(\d+(?:[,\-–—]\d+)*)\]", txt):
                    if "【" in txt or _is_non_citation_bracket(m_br.group(1)):
                        continue
                    bad_cite_place.append(f"Nature 模式出现方括号引用而非上标: {m_br.group(0)!r} in {txt[:40]!r}")
            else:  # gb7714
                for m in re.finditer(r"\[(\d+(?:[,，\-–—]\d+)*)\]", txt):
                    if _is_non_citation_bracket(m.group(1)):
                        continue
                    nums = _expand_num_spec(m.group(1))
                    if nums:
                        cite_groups += 1
                        cited_nums.update(nums)

        ref1_status = "FAIL" if bad_cite_place else ("WARN" if (cite_groups == 0 and ref_idx is None) else "PASS")
        ref1_summary = (
            f"正文识别 {cite_groups} 处引用组 ({citation_style})"
            if ref1_status != "WARN"
            else "正文未检出可识别引文且缺少标准参考文献标题，无法核验"
        )
        results.append(
            CheckResult(
                "Ref-1",
                f"引用规范-1({'Nature上标置标点前' if citation_style == 'nature' else 'GB/T 7714方括号引用'})",
                ref1_status,
                ref1_summary,
                bad_cite_place[:10],
            )
        )

        # Bibliography block
        ref_paras: list[ET.Element] = []
        if ref_idx is not None:
            end_r = next(
                (
                    i
                    for i in range(ref_idx + 1, len(blocks))
                    if blocks[i].tag == f"{W}p"
                    and not _is_toc_para(blocks[i])
                    and (
                        _p_style(blocks[i]) in CH1_STYLES
                        or re.match(r"^(?:个人简历|作者简历|作者简介|附录|致谢)", _p_text(blocks[i]).strip().replace("\u3000", "").replace(" ", ""))
                    )
                ),
                len(blocks),
            )
            ref_paras = [el for el in blocks[ref_idx + 1 : end_r] if el.tag == f"{W}p" and _p_text(el).strip()]

        en_cite_count = sum('EN.CITE' in ''.join(el.text or '' for el in p.iter(W + 'instrText')).upper()
                            for root in parsed.values() for p in root.iter(W + 'p'))
        en_cite_count += sum('EN.CITE' in el.get(W + 'instr', '').upper()
                             for root in parsed.values() for el in root.iter(W + 'fldSimple'))

        ref_nums: list[int] = []
        ref_fmt_issues: list[str] = []
        ref_style_issues: list[str] = []
        seen_titles: dict[str, int] = {}
        dup_refs: list[str] = []
        cjk_seen_in_refs = False
        ref_lang_order_issues: list[str] = []

        if en_cite_count > 0 and not template and not preset.get('audit', {}).get('allow_live_citations', False):
            ref_fmt_issues.append(f"文档仍含 {en_cite_count} 个未解除域代码的 EN.CITE 标记")

        for rp in ref_paras:
            rt = _p_text(rp).strip()
            m_n = re.match(r"^(?:\[(\d+)\]|(\d+)\.)\s*(.+)$", rt)
            if not m_n:
                ref_fmt_issues.append(f"文献条目缺规范编号: {rt[:45]!r}")
                continue
            rnum = int(m_n.group(1) or m_n.group(2))
            rbody = m_n.group(3)
            ref_nums.append(rnum)

            # Check paragraph line spacing (16 pt = 320 dxa exact) & justified
            sp = styles.effective_spacing(rp)
            jc = styles.effective_jc(rp)
            ref_line = str(round(20 * float(typography.get('ref_line_pt', 16))))
            if sp.get("line") != ref_line or sp.get("lineRule") != "exact" or jc != "both":
                ref_style_issues.append(f"文献 {rnum}: 行距={sp.get('line')}/{sp.get('lineRule')} 对齐={jc} (应为 exact {ref_line} / both)")

            # Duplicate check
            norm_key = re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", rbody.lower())[:120]
            if norm_key in seen_titles:
                dup_refs.append(f"文献 {rnum} 疑似与文献 {seen_titles[norm_key]} 重复")
            else:
                seen_titles[norm_key] = rnum

            is_cjk_ref = bool(re.search(r"[\u4e00-\u9fff]", rbody))
            if is_cjk_ref:
                cjk_seen_in_refs = True
            elif cjk_seen_in_refs and citation_style == "nature":
                ref_lang_order_issues.append(f"文献 {rnum}: 中文文献之后又出现英文文献（应先英后中）")

            if citation_style == "nature":
                stripped_doi = DOI_RE.sub("", rbody).strip()
                is_published_vol_page = bool(
                    PUBLISHED_VOL_PAGE_RE.search(stripped_doi)
                    or (re.search(r"\(\d{4}\)", stripped_doi) and re.search(r"\d+\s*[–-]\s*\d+", stripped_doi))
                )
                is_ahead_online = bool(DOI_RE.search(rbody) and not is_published_vol_page)
                if not is_cjk_ref:
                    has_italic = any(styles.effective_italic(r, rp) for r in rp.findall(f"./{W}r") if (_r_text(r) or "").strip())
                    has_bold = any(styles.effective_bold(r, rp) for r in rp.findall(f"./{W}r") if (_r_text(r) or "").strip())
                    if not has_italic or (not is_ahead_online and not has_bold):
                        ref_fmt_issues.append(f"文献 {rnum} Nature 格式缺少刊名斜体({has_italic})或卷号加粗({has_bold})")
                    runs = rp.findall(f'./{W}r')
                    ref_fmt_issues.extend(f'文献 {rnum}: {issue}' for issue in _nature_author_title(runs, rp, styles))
                    if re.search(r"\b\d+-\d+\s*\(\d{4}\)", rbody):
                        ref_fmt_issues.append(f"文献 {rnum} 页码使用了短连字符 - 而不是 en-dash –")
                    if re.search(r"\(\d{4}\)\.$", rbody):
                        ref_fmt_issues.append(f"文献 {rnum} 年份括号后不应有多余句点")
                if preset.get("audit", {}).get("forbid_published_doi", True) and DOI_RE.search(rbody) and not is_ahead_online:
                    ref_fmt_issues.append(f"文献 {rnum} 已正式见刊（含卷期页码），除在线未见刊外请移除 DOI")
            else:  # gb7714
                marker = re.search(r"\[(?:[A-Z]{1,2})(?:/(?:OL|CD|DK|MT))?\]", rbody)
                if not marker:
                    ref_fmt_issues.append(f"文献 {rnum} GB/T 7714 缺少文献类型标识 [J]/[M]/[D] 等")
                elif marker.group() in ('[J]', '[J/OL]'):
                    tail = rbody[marker.end():].strip(' .，,。')
                    if not re.search(r'[^\d\W][^\n]{1,},\s*(?:18|19|20)\d{2}\b', tail, re.UNICODE):
                        ref_fmt_issues.append(f'文献 {rnum}: 期刊条目缺少可识别的刊名或出版年')

        ref_seq_issues: list[str] = []
        if ref_nums and ref_nums != list(range(1, len(ref_nums) + 1)):
            ref_seq_issues.append(f"参考文献编号不连续: 数量={len(ref_nums)}, 最大={max(ref_nums)}")
        expected_refs = set(ref_nums)
        ref_closure_issues: list[str] = []
        if expected_refs - cited_nums:
            ref_closure_issues.append(f"文末存在未被正文引用的文献: {sorted(expected_refs - cited_nums)}")
        if cited_nums - expected_refs:
            ref_closure_issues.append(f"正文引用序号越界或文末缺录: {sorted(cited_nums - expected_refs)}")

        ref_empty = not ref_nums and not cited_nums and en_cite_count == 0
        ref2_has_fail = bool(ref_seq_issues or ref_fmt_issues or ref_lang_order_issues or dup_refs)
        results.append(
            CheckResult(
                "Ref-2-7",
                f"引用规范-2..7({citation_style}格式/去DOI/序/去重)",
                "FAIL" if ref2_has_fail else ("WARN" if ref_empty else "PASS"),
                "未检出文献条目或正文引文，无法核验文献格式" if (ref_empty and not ref2_has_fail) else f"共核验 {len(ref_nums)} 条文献 (EN.CITE={en_cite_count})",
                (ref_seq_issues + ref_lang_order_issues + dup_refs + ref_fmt_issues)[:15]
                or (["未找到标准参考文献节且正文未检出可识别引文"] if ref_empty else []),
            )
        )
        results.append(
            CheckResult(
                "Ref-8",
                "引用规范-8(正文与参考文献双向闭环)",
                "FAIL" if ref_closure_issues else ("WARN" if ref_empty else "PASS"),
                "未检出文献或引文，无法验证双向闭环" if (ref_empty and not ref_closure_issues) else ("全部文献引用双向闭环" if not ref_closure_issues else "存在未引或越界文献"),
                ref_closure_issues or (["缺少参考文献或正文引文条目"] if ref_empty else []),
            )
        )
        results.append(
            CheckResult(
                "Ref-9",
                "引用规范-9(参考文献固定16磅两端对齐)",
                "FAIL" if ref_style_issues else ("WARN" if ref_empty else "PASS"),
                "未检出文献段落，无法核验文献行距与对齐" if (ref_empty and not ref_style_issues) else ("全部文献段落符合固定 16 磅与两端对齐" if not ref_style_issues else f"发现 {len(ref_style_issues)} 处行距/对齐偏差"),
                ref_style_issues[:10] or (["缺少可核验的参考文献段落"] if ref_empty else []),
            )
        )

        # 6. First-person pronouns & consistent self-term (Acad-14; exclude Acknowledgements, legal declarations, and CV entries)
        first_person_hits: list[str] = []
        self_terms: dict[str, int] = {}
        legal_skip = (
            "本人郑重声明", "本人声明", "本人完全了解", "本人了解学校", "本人承诺",
            "授权学校保存", "学位论文使用授权", "未经作者授权", "版权与使用声明", "论文作者签名",
        )
        ack_candidates = [
            i for i, el in enumerate(blocks)
            if el.tag == f"{W}p"
            and not _is_toc_para(el)
            and re.fullmatch(r"致[\s\u3000]*谢", _p_text(el).strip())
        ]
        ack_heading = ack_candidates[-1] if ack_candidates else None
        ack_paras = set()
        if ack_heading is not None:
            ack_end = next(
                (
                    j for j in range(ack_heading + 1, len(blocks))
                    if blocks[j].tag == f"{W}p"
                    and (
                        "独创性" in _p_text(blocks[j])
                        or _p_text(blocks[j]).strip().replace("\u3000", "").replace(" ", "").startswith(("个人简历", "作者简历", "附录", "在学期间", "版权"))
                        or (_p_style(blocks[j]) in CH1_STYLES and not re.fullmatch(r"致[\s\u3000]*谢", _p_text(blocks[j]).strip()))
                    )
                ),
                len(blocks),
            )
            ack_paras = {id(p) for el in blocks[ack_heading:ack_end] for p in ([el] if el.tag == f"{W}p" else list(el.iter(f"{W}p")))}
        cv_heading_idx = next(
            (
                i
                for i, el in enumerate(blocks)
                if el.tag == f"{W}p"
                and not _is_toc_para(el)
                and re.match(r"^(?:个人简历|作者简历|作者简介)", _p_text(el).strip().replace("\u3000", "").replace(" ", ""))
            ),
            None,
        )
        cv_paras_list: list[ET.Element] = []
        cv_paras_ids: set[int] = set()
        if cv_heading_idx is not None:
            cv_end_idx = next(
                (
                    j
                    for j in range(cv_heading_idx + 1, len(blocks))
                    if blocks[j].tag == f"{W}p"
                    and not _is_toc_para(blocks[j])
                    and (
                        _p_style(blocks[j]) in CH1_STYLES
                        or re.match(r"^(?:附录|致谢|学位论文独创性)", _p_text(blocks[j]).strip().replace("\u3000", "").replace(" ", ""))
                    )
                ),
                len(blocks),
            )
            cv_paras_list = [
                p
                for el in blocks[cv_heading_idx + 1 : cv_end_idx]
                for p in ([el] if el.tag == f"{W}p" else list(el.iter(f"{W}p")))
                if _p_text(p).strip()
            ]
            cv_paras_ids = {id(p) for p in cv_paras_list}
        for p in all_paras:
            if id(p) in ack_paras or id(p) in cv_paras_ids:
                continue
            ps = _p_style(p)
            t = _p_text(p).strip()
            if not t or ps in BIB_STYLES or any(k in t for k in legal_skip) or "不得出现“我" in t or "我如何如何" in t:
                continue
            for m_fp in FIRST_PERSON_RE.finditer(t):
                first_person_hits.append(f"{m_fp.group(0)}@{t[:36]}")
            for term in ("本论文", "本研究", "本工作", "本课题"):
                cnt = len(re.findall(term, t))
                if cnt:
                    self_terms[term] = self_terms.get(term, 0) + cnt
        self_status = "PASS" if (not first_person_hits and len(self_terms) <= 1) else "FAIL"
        results.append(CheckResult("Acad-14", "学术规范-14(禁第一人称/自称一致)", self_status, f"第一人称={len(first_person_hits)}, 自称分布={self_terms}", first_person_hits[:8]))

        # Author CV requirement & blind redaction check (Acad-CV)
        want_cv = bool(preset.get("modes", {}).get(mode, {}).get("include_cv", mode in ("phd-thesis", "master-thesis", "proposal", "midterm")))
        cv_issues: list[str] = []
        cv_blind_leaks: list[str] = []
        if args.review_stage == "blind":
            for cvp in cv_paras_list:
                cvt = _p_text(cvp).strip()
                if cvt.startswith("姓名：") and "匿名" not in cvt and "隐去" not in cvt:
                    cv_blind_leaks.append(f"盲审作者简历仍含未脱敏基本信息: {cvt[:36]!r}")
                elif re.match(r"^(?:19|20)\d{2}", cvt) and "匿名" not in cvt and "隐去" not in cvt:
                    cv_blind_leaks.append(f"盲审作者简历仍含未脱敏教育/工作经历: {cvt[:36]!r}")
                elif (re.match(r"^\d+\.\s+", cvt) or "et al." in cvt) and "匿名" not in cvt and "隐去" not in cvt and "Fictional" not in cvt:
                    cv_blind_leaks.append(f"盲审作者简历成果列表未隐去署名: {cvt[:36]!r}")
            cv_issues.extend(cv_blind_leaks)
            cv_status = "FAIL" if cv_issues else "PASS"
            cv_summary = (
                "作者简历符合盲审匿名要求"
                if not cv_issues
                else f"作者简历存在 {len(cv_issues)} 处盲审脱敏遗漏"
            )
        elif want_cv:
            if cv_heading_idx is None:
                cv_issues.append("参考文献后缺少“个人简历/作者简历及在学期间发表的学术论文与研究成果”章节")
            elif ref_idx is not None and cv_heading_idx <= ref_idx:
                cv_issues.append("“个人简历/作者简历”应置于参考文献之后")
            elif not cv_paras_list:
                cv_issues.append("“个人简历/作者简历”标题下缺少简历或成果列表正文")
            cv_status = "FAIL" if cv_issues else "PASS"
            cv_summary = f"已检测作者简历（{len(cv_paras_list)} 段）" if not cv_issues else cv_issues[0]
        else:
            cv_status = "INFO"
            cv_summary = f"当前模式 ({mode}) 不强制要求作者简历"
        results.append(CheckResult("Acad-CV", "学术规范-21(作者简历与成果列表)", cv_status, cv_summary, cv_issues[:8]))

        # Units, subscripts, experimental time units, punctuation, Chinese spaces, enzymes, genes/species, abbreviations
        unit_issues: list[str] = []
        formula_review: list[str] = []
        punct_issues: list[str] = []
        space_issues: list[str] = []
        bio_italic_issues: list[str] = []
        user_genes = [g.strip() for g in (args.genes or "").split(",") if g.strip()]
        user_species = [s.strip() for s in (args.species or "").split(",") if s.strip()] + list(BUILTIN_SPECIES)

        for p in prose_paras:
            ps = _p_style(p)
            t = _p_text(p).strip()
            if not t:
                continue
            for m in BAD_UNIT_U_RE.finditer(t):
                unit_issues.append(f"单位 u 应改 µ/μ: {m.group(0)!r}")
            for m in NO_SPACE_UNIT_RE.finditer(t):
                unit_issues.append(f"数字与单位间缺空格: {m.group(0)!r}")
            if FORMULA_RE.search(t):
                formula_review.extend(_formula_review(p, styles))
            if ps in (BODY_STYLES | ABS_ZH_STYLES | FIG_CAP_STYLES):
                for m in SUBPANEL_BAD_SEP_RE.finditer(t):
                    punct_issues.append(f"多子图图序请统一用中文逗号“，”分隔: {m.group(0)}")
                for m in CJK_ASCII_PUNCT_RE.finditer(t):
                    punct_issues.append(f"中文行文混用英文标点: {m.group(0)!r}")
                for m_p in PAREN_ZH_CONTENT_RE.finditer(t):
                    inner = m_p.group(1)
                    if "," not in inner:
                        continue
                    if (
                        (re.search(r"[\u4e00-\u9fff]\s*,|,\s*[\u4e00-\u9fff]", inner) or VENDOR_KEYWORDS_RE.search(inner))
                        and not re.search(r"et al\.\s*,\s*\d{4}", inner)
                    ):
                        punct_issues.append(f"中文括号/设备厂商名间误用英文逗号（应为中文逗号“，”）: {m_p.group(0)[:40]!r}")
                for m_v in PAREN_VENDOR_DUNHAO_RE.finditer(t):
                    inner_v = m_v.group(1)
                    parts_v = [seg.strip() for seg in inner_v.split("、")]
                    if len(parts_v) == 2 and not re.search(r"等(?:试剂|仪器|生物|设备|公司|厂商|供应商|耗材|品牌|$)", inner_v):
                        if bool(VENDOR_BRAND_RE.search(parts_v[0])) ^ bool(VENDOR_BRAND_RE.search(parts_v[1])):
                            punct_issues.append(f"设备/试剂与厂商间误用顿号（应为中文逗号“，”）: {m_v.group(0)[:40]!r}")
            if ps in ABS_EN_STYLES and re.search(r"[，。；：（）“”‘’、]", t):
                punct_issues.append(f"英文摘要含中文标点: {t[:40]!r}")
            for m in ENZYME_NO_SPACE_RE.finditer(t):
                bio_italic_issues.append(f"内切酶罗马数字前缺空格: {m.group(0)!r}")
            for r in p.findall(f".//{W}r"):
                rt = _r_text(r)
                if not rt:
                    continue
                is_it = styles.effective_italic(r, p)
                for g in user_genes:
                    if re.search(rf"\b{re.escape(g)}\b", rt) and not is_it:
                        bio_italic_issues.append(f"基因未斜体: {g} in {rt[:30]!r}")
                for sp in user_species:
                    if sp in rt and not is_it:
                        bio_italic_issues.append(f"拉丁物种名未斜体: {sp} in {rt[:30]!r}")

        space_scope_paras = [
            p
            for p in blocks
            if p.tag == f"{W}p"
            and not _is_toc_para(p)
            and _p_style(p) in (BODY_STYLES | ABS_ZH_STYLES | ABS_EN_STYLES | FIG_CAP_STYLES | TBL_NOTE_STYLES)
            and _p_style(p) not in (CH1_STYLES | H2_STYLES | H3_STYLES | H4_STYLES)
        ]
        for p in space_scope_paras:
            t = _live_p_text(p).strip()
            if not t or any(k in t for k in legal_skip) or "____" in t or _p_style(p) in BIB_STYLES:
                continue
            for m_sp in CJK_HALF_SPACE_RE.finditer(t):
                space_issues.append(f"中文字符后留空使用了英文半角空格（请改用全角空格“　”或删除多余空格）: {m_sp.group(0)!r} in {t[:36]!r}")
            for m_fs in ASCII_FULL_SPACE_RE.finditer(t):
                space_issues.append(f"英文/数字后误用中文全角空格（应使用英文半角空格）: {m_fs.group(0)!r} in {t[:36]!r}")

        unit_status = 'FAIL' if unit_issues else 'WARN' if formula_review else 'PASS'
        unit_summary = (f'检出 {len(unit_issues)} 处单位问题，{len(formula_review)} 处化学式待复核'
                        if unit_issues or formula_review else '单位、微升符号及化学式规范')
        results.append(CheckResult("Err-4-5-6", "常见错误-4/5/6(化学式下标/µL/数字单位空格)",
                                   unit_status, unit_summary, (unit_issues + formula_review)[:12]))

        ch1_headings = [
            i for i, el in enumerate(blocks)
            if el.tag == f"{W}p"
            and not _is_toc_para(el)
            and _p_style(el) in CH1_STYLES
            and not _is_front_title(_p_text(el))
            and (_para_num_info(el, styles, numbering_defs) is not None or re.match(r"^(?:第[一二三四五六七八九十\d]+章|Chapter\s+\d+)\b", _p_text(el).strip(), re.I))
        ]
        methods_range: tuple[int, int] | None = None
        for idx_pos, h_idx in enumerate(ch1_headings):
            htxt = _p_text(blocks[h_idx]).strip()
            if re.search(r"^(?:第二章|Chapter\s+2\b)", htxt, re.I) or re.search(r"(?:材料.*方法|实验方法|实验程序|研究方法|方法和步骤|Materials?\s+and\s+Methods?)", htxt, re.I):
                next_h = ch1_headings[idx_pos + 1] if idx_pos + 1 < len(ch1_headings) else body_end
                methods_range = (h_idx, next_h)
                break
        has_cjk_t, has_eng_t = False, False
        if methods_range is not None:
            exp_txt = "\n".join(_p_text(el) for el in blocks[methods_range[0] : methods_range[1]])
            exp_clean = FIG_TBL_REF_STRIP_RE.sub("", exp_txt)
            has_cjk_t = bool(CJK_TIME_RE.search(exp_clean))
            has_eng_t = bool(ENG_TIME_WORD_RE.search(exp_clean) or ENG_TIME_SYM_RE.search(exp_clean))
        results.append(CheckResult("Err-7", "常见错误-7(实验程序时间单位中英文统一)", "FAIL" if (has_cjk_t and has_eng_t) else "PASS", f"实验程序中文时间={has_cjk_t}, 英文时间={has_eng_t}", ["实验程序章节同时混用中文时间单位与英文缩写（须全文统一中文或英文）"] if (has_cjk_t and has_eng_t) else []))
        results.append(CheckResult("Err-2-11-12", "常见错误-2/11/12(中英文标点与多子图/厂商逗号)", "PASS" if not punct_issues else "FAIL", "中英文标点、多子图及设备厂商逗号规范" if not punct_issues else f"发现 {len(punct_issues)} 处标点混用", punct_issues[:12]))
        results.append(CheckResult("Err-14", "常见错误-14(中文后全角空格与英文后半角空格)", "PASS" if not space_issues else "FAIL", "中英文空格规范" if not space_issues else f"发现 {len(space_issues)} 处中英文空格异常", space_issues[:12]))
        results.append(CheckResult("Acad-15-16", "学术规范-15/16/常见错误-13(基因/物种/内切酶斜体与空格)", "PASS" if not bio_italic_issues else "FAIL", "物种名、基因名及内切酶检查通过" if not bio_italic_issues else f"发现 {len(bio_italic_issues)} 处斜体或空格问题", bio_italic_issues[:12]))

        # Layout-7 abbreviation first expansion check per scope
        abs_zh = [p for p in all_paras if _p_style(p) in ABS_ZH_STYLES]
        abs_en = [p for p in all_paras if _p_style(p) in ABS_EN_STYLES]
        dup_abbrs: list[str] = []
        for scope_name, plist in (("中文摘要", abs_zh), ("英文摘要", abs_en), ("正文", main_paras)):
            seen_ab: dict[str, int] = {}
            for p in plist:
                for m in ABBR_PAREN_RE.finditer(_p_text(p)):
                    ab = m.group(3)
                    if ab in {"et", "al", "vs", "SD", "SE"} or "/" in ab or ab.isdigit():
                        continue
                    seen_ab[ab] = seen_ab.get(ab, 0) + 1
            for ab, cnt in seen_ab.items():
                if cnt > 1:
                    dup_abbrs.append(f"{scope_name}内缩写 {ab} 重复展开 {cnt} 次")
        results.append(CheckResult("Layout-7", "排版-7(摘要与正文首次展开不重复)", "PASS" if not dup_abbrs else "WARN", "未发现同域内重复展开的缩写" if not dup_abbrs else f"疑似重复展开 {len(dup_abbrs)} 项", dup_abbrs[:10]))

        # Blind review check
        if args.review_stage == "blind":
            tokens = [s.strip() for s in (args.sensitive_names or '').split(',') if s.strip()]
            blind_issues = [f'发现指定身份词：{part}' for part in identity_parts(parsed, tokens)]
            blind_issues.extend(cv_blind_leaks)
            ack_candidates = [
                i for i, el in enumerate(blocks)
                if el.tag == f"{W}p"
                and not (_p_style(el) or "").lower().startswith("toc")
                and re.fullmatch(r"致[\s\u3000]*谢", _p_text(el).strip())
            ]
            ack_i = ack_candidates[-1] if ack_candidates else None
            if ack_i is not None:
                ack_txt = "\n".join(_p_text(el) for el in blocks[ack_i + 1 :] if el.tag == f"{W}p" and "独创性" not in _p_text(el) and "声明" not in _p_text(el))
                if "匿名评审" not in ack_txt and len(ack_txt.strip()) > 30:
                    blind_issues.append("盲审版本致谢段含非匿名正文，请核实是否已隐去导师与姓名")
            status = 'FAIL' if blind_issues else 'PASS' if tokens else 'WARN'
            summary = {'FAIL': f'发现 {len(blind_issues)} 处署名泄露隐患',
                       'PASS': f'已核查 {len(tokens)} 个指定词在 XML/关系部件中的残留',
                       'WARN': '未提供敏感词，无法核验已知姓名/学号是否残留'}[status]
            results.append(CheckResult('Acad-10-Blind', '盲审身份字段核查', status, summary, blind_issues[:10]))

        # Word count
        body_prose = "\n".join(
            _p_text(p)
            for p in main_paras
            if _p_style(p) not in (FIG_CAP_STYLES | TBL_TITLE_STYLES | TBL_NOTE_STYLES)
            and not re.match(r"^(?:图|补充图|表|附表)\s*(?:[1-9]\d*|[A-D])", _p_text(p).strip())
        )
        zh_c = len(re.findall(r"[\u4e00-\u9fff]", body_prose))
        en_w = len(re.findall(r"[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*", body_prose))
        total_w = zh_c + en_w
        stats["body_zh_chars"] = zh_c
        stats["body_en_words"] = en_w
        stats["body_total_words_excl_refs_captions"] = total_w
        wc_status = "INFO"
        wc_msg = f"正文（不含图注与参考文献）计约 {total_w} 字（中文 {zh_c} + 英文/数字词 {en_w}）"
        wc_range = preset.get("modes", {}).get(mode, {}).get("word_count_range")
        if isinstance(wc_range, list) and len(wc_range) == 2 and not template:
            lo, hi = wc_range[0], wc_range[1]
            if (lo is not None and total_w < lo) or (hi is not None and total_w > hi):
                wc_status = "WARN"
                bound_str = f"{lo:,}–{hi:,}" if (lo is not None and hi is not None) else f">= {lo:,}"
                wc_msg += f" — 当前模式建议 {bound_str} 字"
        results.append(CheckResult("WordCount", "正文字数统计", wc_status, wc_msg, []))

        custom_style = False
        if args.profile_config:
            override = json.loads(pathlib.Path(args.profile_config).read_text(encoding='utf-8'))
            if not isinstance(override, dict):
                raise ValueError('--profile-config must contain a JSON object')
            cs_val = override.get('citation_style')
            if cs_val is not None and cs_val not in ('nature', 'gb7714'):
                raise ValueError(f"Unsupported profile citation_style {cs_val!r}; expected 'nature' or 'gb7714'")
            custom_style = cs_val in ('nature', 'gb7714')
        _use_audit_policy(results, preset, template, args.citation_style != 'auto' or custom_style
                          or args.preset == 'cn-gb7714-general')

    return stats, results


def _report_paths(args: argparse.Namespace) -> dict[str, pathlib.Path]:
    paths = {}
    inputs = [pathlib.Path(value) for value in (args.docx, args.base_docx, args.template, args.profile_config) if value]
    for kind, value in (('json', args.json_report), ('md', args.md_report)):
        if not value:
            continue
        path = pathlib.Path(value).resolve()
        if path == SKILL_DIR or SKILL_DIR in path.parents:
            raise PermissionError(f'Refusing to write reports inside skill directory: {path}')
        if any(same_path(path, source) for source in [*inputs, *paths.values()]):
            raise ValueError(f'Report path aliases an input or other report: {path}')
        if path.suffix.lower() != '.' + kind:
            raise ValueError(f'Report must have a .{kind} extension: {path}')
        paths[kind] = path
    return paths


def _manual_review(review_stage: str) -> list[dict[str, str]]:
    tasks = [
        ("rendered-pages", "按学校要求，在 Word 打印布局或最终 PDF 中核对封面、各节起始页、页码与页眉，并检查空白页与孤行。"),
        ("rendered-figures-tables", "如有图表，在最终 PDF 中逐一检查图注与图片是否同页、图中文字是否不小于 6 号字（7.5 pt）且标尺统一、表格跨页表头是否重复。"),
        ("refresh-fields", "如有目录或交叉引用，在 Word 副本中更新这些域，重新导出 PDF，抽查目录页码、图表编号和跳转位置。"),
    ]
    if review_stage == "blind":
        tasks.append(("blind-images-signatures", "逐页查看图片、扫描签名和图中标识，并打开嵌入附件，确认作者和导师身份已按盲审要求处理。"))
    return [{"id": key, "status": "PENDING", "action": action} for key, action in tasks]


def _write_reports(paths, stats: dict, results: list[CheckResult], counts: dict[str, int],
                   manual_review: list[dict[str, str]]) -> None:
    notice = "完整报告可能包含本地文件路径和论文原文片段。向他人分享前请检查这些内容。"
    payload = {"stats": stats, "checks": [asdict(r) for r in results], "review_summary": counts,
               "manual_review": manual_review, "report_notice": notice}
    for kind, destination in paths.items():
        if kind == 'json':
            text = json.dumps(payload, ensure_ascii=False, indent=2) + '\n'
        else:
            lines = ['# Bio-Thesis Formatting Audit Report', '', f"- DOCX: `{stats.get('docx')}`",
                     f"- Template: `{stats.get('template') or 'preset'}`", f"- SHA-256: `{stats.get('sha256')}`",
                     f"- Mode: `{stats.get('detected_mode')}` | Preset: `{stats.get('preset')}` | Citation: `{stats.get('citation_style')}`",
                     '', notice, '', '## 自动检查统计', '',
                     ' · '.join(f'{status} {count}' for status, count in counts.items()),
                     '', '## 人工待复核', '']
            lines.extend(f"- [ ] {item['action']}" for item in manual_review)
            lines.extend(['', '## Findings', ''])
            for result in results:
                lines.append(f'- **[{result.status}] {result.check_id}** ({result.checklist_ref}): {result.summary}')
                lines.extend(f'  - {detail}' for detail in result.details)
            text = '\n'.join(lines) + '\n'
        with staged_file(destination) as stage:
            stage.write_text(text, encoding='utf-8')


def add_arguments(p: argparse.ArgumentParser) -> None:
    p.add_argument("--docx", required=True, help="Target .docx to audit (read-only)")
    p.add_argument("--base-docx", help="Optional baseline .docx to verify media byte identity")
    p.add_argument("--template", help="Compare layout and style definitions with a supplied .docx or .dotx")
    p.add_argument("--mode", choices=("auto", "phd-thesis", "master-thesis", "proposal", "midterm", "rotation"), default="auto")
    p.add_argument("--review-stage", choices=("final", "blind"), default="final")
    p.add_argument("--preset", choices=("cn-lifesci-default", "cn-gb7714-general"), default="cn-lifesci-default")
    p.add_argument("--profile-config", help="Optional custom university/department JSON override")
    p.add_argument("--citation-style", choices=("auto", "nature", "gb7714"), default="auto")
    p.add_argument("--genes", help="Comma-separated gene symbols requiring italics")
    p.add_argument("--species", help="Comma-separated Latin species binomials requiring italics")
    p.add_argument("--sensitive-names", help="Comma-separated author/advisor tokens to scan in blind mode")
    p.add_argument("--fail-on-warn", action="store_true", help="Also exit 1 for audit warnings; INFO is advisory")
    p.add_argument("--json-report", help="Optional JSON report path (outside skill dir)")
    p.add_argument("--md-report", help="Optional Markdown report path (outside skill dir)")


def run_parsed(args: argparse.Namespace) -> int:
    try:
        paths = _report_paths(args)
        stats, results = run_audit(args)
        counts = dict.fromkeys(("PASS", "WARN", "FAIL", "INFO"), 0)
        for result in results:
            counts[result.status] += 1
        _write_reports(paths, stats, results, counts, _manual_review(args.review_stage))
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    for r in results:
        print(f"[{r.status}] {r.check_id} ({r.checklist_ref}): {r.summary}")
        for d in r.details[:6]:
            print(f"    -> {d}")
    print(f"AUDIT_SUMMARY=FAIL:{counts['FAIL']} WARN:{counts['WARN']} PASS:{counts['PASS']}")
    return 1 if counts["FAIL"] or (args.fail_on_warn and counts["WARN"]) else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only OOXML & 67-item compliance audit for theses and reports.")
    add_arguments(parser)
    args = parser.parse_args(argv)
    return run_parsed(args)


if __name__ == "__main__":
    raise SystemExit(main())
