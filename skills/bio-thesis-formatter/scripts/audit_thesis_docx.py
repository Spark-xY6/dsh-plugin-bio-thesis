#!/usr/bin/env python3
"""Read-only OOXML & 65-item compliance auditor using Python stdlib only."""

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

from docx_package import identity_parts, read_template, same_path, staged_file
from init_thesis_docx import _resolve_page_dxa, load_effective_preset

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent.resolve()
SCHOOL_CHECKS = {
    'Err-9', 'Layout-9', 'Err-8', 'Layout-BodySpacing', 'Ref-9', 'Tbl-3Line',
    'Fig-1', 'Fig-3-8', 'Fig-4', 'Fig-9-10', 'Acad-2', 'Acad-14', 'Err-2-11-12',
    'Err-7', 'Layout-7', 'WordCount',
}

CH1_STYLES = {"Heading1Chapter", "04", "Heading1", "heading 1"}
H2_STYLES = {"Heading2Sec", "05XX", "Heading2", "heading 2"}
H3_STYLES = {"Heading3Subsec", "06XXX", "Heading3", "heading 3"}
H4_STYLES = {"Heading4Para", "07XXXX", "Heading4", "heading 4"}
ABS_ZH_STYLES = {"AbstractBodyZH", "02"}
ABS_EN_STYLES = {"AbstractBodyEN", "03"}
BODY_STYLES = {"BodyTextZH", "08", "Normal", "a"}
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
CJK_TIME_RE = re.compile(r"(?:\d|\s)(?:小时|分钟|秒)(?:后|内|时|，|。|、|）|$)")
ENG_TIME_RE = re.compile(r"\b\d+(?:\.\d+)?\s*(?:hr|hrs|min|mins|sec|secs)\b", re.IGNORECASE)
SUBPANEL_DUNHAO_RE = re.compile(r"图\s*[A-D1-9]\d*\.\d+\s*[A-Z](?:、[A-Z])+")
CJK_ASCII_PUNCT_RE = re.compile(r"[\u4e00-\u9fff][,;:\(\)][\u4e00-\u9fff]")
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


def _r_text(r: ET.Element) -> str:
    return "".join(t.text or "" for t in r.iter(f"{W}t"))


def _p_style(p: ET.Element) -> str | None:
    ps = p.find(f"./{W}pPr/{W}pStyle")
    return ps.get(f"{W}val") if ps is not None else None


def _tag_local(el: ET.Element) -> str:
    return el.tag.split("}", 1)[-1]


def _on(node: ET.Element | None) -> bool:
    return node is not None and node.get(W + 'val', '1') not in ('0', 'false', 'off')


def _section_geometry(section: ET.Element) -> dict[str, int | str | None]:
    geometry = {}
    for tag, keys in (('pgSz', ('w', 'h')), ('pgMar', ('top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'))):
        node = section.find(W + tag)
        for key in keys:
            value = node.get(W + key) if node is not None else None
            geometry[key] = int(value) if value is not None else None
        if tag == 'pgSz':
            geometry['orient'] = node.get(W + 'orient', 'portrait') if node is not None else None
    return geometry


def _instructions(root: ET.Element, command: str) -> bool:
    pattern = re.compile(r'^\s*' + re.escape(command) + r'(?:\s|$)', re.IGNORECASE)
    if any(pattern.search(field.get(W + 'instr', '')) for field in root.iter(W + 'fldSimple')):
        return True
    for paragraph in root.iter(W + 'p'):
        text = ' '.join(node.text or '' for node in paragraph.iter(W + 'instrText'))
        joined = ''.join(node.text or '' for node in paragraph.iter(W + 'instrText'))
        if pattern.search(text) or pattern.search(joined):
            return True
    return False


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
                target = rels_map.get(reference.get(R + 'id', ''), '').lstrip('/')
                part = target if target.startswith('word/') else 'word/' + target
                inherited[reference.get(W + 'type', 'default')] = parsed.get(part)
        active = tuple(sorted(role for role, root in footers.items()
                              if root is not None and _instructions(root, 'PAGE')))
        header_roles = tuple(sorted(role for role, root in headers.items() if root is not None))
        result.append((fmt, start, kind, _on(section.find(W + 'titlePg')), active, header_roles))
    return result


def _template_snapshot(template_path: str) -> dict:
    items, _ = read_template(pathlib.Path(template_path))
    parts = {k: ET.fromstring(v) for k, v in items.items() if k.endswith(('.xml', '.rels'))}
    document = parts['word/document.xml']
    rels = parts.get('word/_rels/document.xml.rels')
    rels_map = {r.get('Id', ''): r.get('Target', '').lstrip('/') for r in rels} if rels is not None else {}
    sections = list(document.iter(W + 'sectPr'))
    return {
        'settings': parts.get('word/settings.xml'),
        'geometry': [_section_geometry(sec) for sec in sections],
        'numbering': _section_numbering(sections, parts, rels_map),
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


def _border_visible(border) -> bool:
    return border is not None and border.get(W + 'val', 'nil') not in ('nil', 'none')


def _matches_rule(border, size: str) -> bool:
    return border is not None and border.get(W + 'val') == 'single' and border.get(W + 'sz') == size


def _table_problems(table: ET.Element) -> tuple[list[str], bool]:
    props = table.find(W + 'tblPr')
    borders = table.find(f'./{W}tblPr/{W}tblBorders')
    if borders is None:
        inherited = props is not None and props.find(W + 'tblStyle') is not None
        return ['无显式边框，须验证表格样式解析结果' if inherited else '缺少显式表格三线边框'], inherited
    problems = []
    for side in ('top', 'bottom'):
        if not _matches_rule(borders.find(W + side), '12'):
            problems.append(f'{side} 应为 1.5 pt 单线')
    for side in ('left', 'right', 'start', 'end', 'insideV', 'insideH'):
        if _border_visible(borders.find(W + side)):
            problems.append(f'{side} 存在额外边框')
    rows = table.findall(W + 'tr')
    if not rows:
        return problems + ['表格缺少可核验的行'], False
    header_count = 0
    for row in rows:
        if not _on(row.find(f'./{W}trPr/{W}tblHeader')):
            break
        header_count += 1
    if not header_count:
        problems.append('首行未标记为可重复表头')
        header_count = 1
    for index, row in enumerate(rows):
        if row.find(f'./{W}tblPrEx/{W}tblBorders') is not None:
            problems.append(f'第 {index + 1} 行存在表级边框例外，须人工确认')
        for cell in row.findall(W + 'tc'):
            cell_borders = cell.find(f'./{W}tcPr/{W}tcBorders')
            if index == header_count - 1:
                bottom = cell_borders.find(W + 'bottom') if cell_borders is not None else None
                if not _matches_rule(bottom, '6'):
                    problems.append('表头分隔线应为 0.75 pt 单线')
            if cell_borders is None:
                continue
            for border in cell_borders:
                if not _border_visible(border):
                    continue
                side = _tag_local(border)
                allowed = (side == 'bottom' and index == header_count - 1
                           or side == 'top' and index == 0 and _matches_rule(border, '12')
                           or side == 'bottom' and index == len(rows) - 1 and _matches_rule(border, '12')
                           or side == 'top' and index == header_count and _matches_rule(border, '6'))
                if not allowed:
                    problems.append(f'第 {index + 1} 行单元格存在额外 {side} 边框')
    return list(dict.fromkeys(problems)), False


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
    def __init__(self, styles_root: ET.Element | None) -> None:
        self.styles: dict[str, ET.Element] = {}
        self.default_fonts: dict[str, str] = {}
        self.default_p_jc: str = "left"
        if styles_root is None:
            return
        rdef = styles_root.find(f"./{W}docDefaults/{W}rPrDefault/{W}rPr/{W}rFonts")
        if rdef is not None:
            for k in ("ascii", "eastAsia", "hAnsi", "cs"):
                v = rdef.get(f"{W}{k}")
                if v:
                    self.default_fonts[k] = v
        for st in styles_root.findall(f"./{W}style"):
            sid = st.get(f"{W}styleId")
            if sid:
                self.styles[sid] = st

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

    def effective_jc(self, p: ET.Element) -> str:
        jc = self.default_p_jc
        for st in self._chain(_p_style(p)):
            el = st.find(f"./{W}pPr/{W}jc")
            if el is not None and el.get(f"{W}val"):
                jc = el.get(f"{W}val", jc)
        el = p.find(f"./{W}pPr/{W}jc")
        if el is not None and el.get(f"{W}val"):
            jc = el.get(f"{W}val", jc)
        return jc

    def effective_spacing(self, p: ET.Element) -> dict[str, str]:
        sp: dict[str, str] = {}
        for st in self._chain(_p_style(p)):
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

    def effective_bold(self, r: ET.Element, p: ET.Element | None = None) -> bool:
        val = False
        if p is not None:
            for st in self._chain(_p_style(p)):
                rpr = st.find(f"./{W}rPr")
                if rpr is not None and rpr.find(f"./{W}b") is not None:
                    b = rpr.find(f"./{W}b")
                    val = b.get(f"{W}val", "1") not in ("0", "false")
        rpr = r.find(f"./{W}rPr")
        if rpr is not None and rpr.find(f"./{W}b") is not None:
            b = rpr.find(f"{W}b")
            val = b.get(f"{W}val", "1") not in ("0", "false")
        return val

    def effective_italic(self, r: ET.Element, p: ET.Element | None = None) -> bool:
        val = False
        if p is not None:
            for st in self._chain(_p_style(p)):
                rpr = st.find(f"./{W}rPr")
                if rpr is not None and rpr.find(f"./{W}i") is not None:
                    i = rpr.find(f"./{W}i")
                    val = i.get(f"{W}val", "1") not in ("0", "false")
        rpr = r.find(f"./{W}rPr")
        if rpr is not None and rpr.find(f"./{W}i") is not None:
            i = rpr.find(f"./{W}i")
            val = i.get(f"{W}val", "1") not in ("0", "false")
        return val

    def effective_ascii_font(self, r: ET.Element, p: ET.Element | None = None) -> str:
        font = self.default_fonts.get("ascii", "Times New Roman")
        if p is not None:
            for st in self._chain(_p_style(p)):
                rf = st.find(f"./{W}rPr/{W}rFonts")
                if rf is not None:
                    font = rf.get(f"{W}ascii") or rf.get(f"{W}hAnsi") or font
        rf = r.find(f"./{W}rPr/{W}rFonts")
        if rf is not None:
            font = rf.get(f"{W}ascii") or rf.get(f"{W}hAnsi") or font
        return font


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


def _extract_runs_skip_toc_fields(p: ET.Element) -> list[ET.Element]:
    out: list[ET.Element] = []
    in_field = False
    for child in list(p):
        tag = _tag_local(child)
        if tag == "hyperlink":
            out.extend(list(child.findall(f"./{W}r")))
        elif tag == "r":
            fc = child.find(f"./{W}fldChar")
            if fc is not None:
                ft = fc.get(f"{W}fldCharType")
                if ft == "begin":
                    in_field = True
                elif ft == "end":
                    in_field = False
            elif not in_field:
                out.append(child)
    return out


def run_audit(args: argparse.Namespace) -> tuple[dict, list[CheckResult]]:
    docx_path = pathlib.Path(args.docx).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    preset = load_effective_preset(args.preset, args.profile_config, None, None, args.citation_style)
    citation_style = preset.get("citation_style", "nature")
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
        if 'word/document.xml' not in parsed or parsed['word/document.xml'].find(W + 'body') is None:
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
        rels_root = parsed.get("word/_rels/document.xml.rels")
        rels_map: dict[str, str] = {}
        if rels_root is not None:
            for rel in list(rels_root):
                rid = rel.get("Id", "")
                tgt = rel.get("Target", "")
                if rel.get("TargetMode") == "External":
                    continue
                rels_map[rid] = tgt
                parts: list[str] = []
                for seg in ("word/" + tgt).split("/"):
                    if seg == "..":
                        if parts:
                            parts.pop()
                    elif seg and seg != ".":
                        parts.append(seg)
                norm = "/".join(parts)
                if norm not in names:
                    broken_rels.append(f"{rid} -> {tgt} ({norm})")
        results.append(
            CheckResult(
                "Pkg-Rels",
                "工程基础-关系闭环",
                "PASS" if not broken_rels else "FAIL",
                "document.xml.rels 内部目标全部存在" if not broken_rels else f"存在 {len(broken_rels)} 个悬空关系",
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

        styles = StyleResolver(parsed.get("word/styles.xml"))
        body = doc_root.find(f"./{W}body")
        assert body is not None
        blocks = list(body)
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
        sect_prs = list(doc_root.iter(f"{W}sectPr"))
        margin_issues: list[str] = []
        if template and len(sect_prs) != len(template['geometry']):
            margin_issues.append(f"分节数 {len(sect_prs)} 与模板 {len(template['geometry'])} 不同")
        expected = template['geometry'] if template else [exp_dims] * len(sect_prs)
        for s_idx, (sp, dimensions) in enumerate(zip(sect_prs, expected), 1):
            for attr, actual in _section_geometry(sp).items():
                want = dimensions[attr]
                if actual == want:
                    continue
                if template or not isinstance(actual, int) or not isinstance(want, int) or abs(actual - want) > 25:
                    unit = '' if attr == 'orient' else ' dxa'
                    margin_issues.append(f"Section {s_idx}: {attr}={actual}{unit} (期望 {want})")
        if not sect_prs:
            margin_issues.append('缺少页面设置分节')

        results.append(
            CheckResult(
                "Layout-1",
                "排版-1(纸张与页边距)",
                "PASS" if (sect_prs and not margin_issues) else "FAIL",
                f"全部 {len(sect_prs)} 节页面设置符合{'模板' if template else '当前预设'}" if not margin_issues else f"发现 {len(margin_issues)} 处页面偏差",
                margin_issues[:10],
            )
        )

        # 3-section & header isolation
        sec_issues: list[str] = []
        numbering = _section_numbering(sect_prs, parsed, rels_map)
        sec_descs = [f'S{i}({fmt}@{start or "续编"}, PAGE={",".join(active) or "-"})'
                     for i, (fmt, start, _, _, active, _) in enumerate(numbering, 1)]
        if template:
            if numbering != template['numbering']:
                sec_issues.append('分节起始方式、首页设置、页眉类型、页码格式或页码域与模板不同')
            actual_styles = _style_properties(parsed.get('word/styles.xml'))
            changed = [key for key, props in template['style_props'].items() if actual_styles.get(key) != props]
            results.append(CheckResult('Layout-TemplateStyles', '模板样式定义', 'WARN' if changed else 'PASS',
                                       f'模板样式属性有 {len(changed)} 项差异' if changed else '模板样式属性匹配', changed[:15]))
        elif len(sect_prs) < 3:
            sec_issues.append(f"仅发现 {len(sect_prs)} 节；规范应至少分 3 节（封面无页码 / 摘要目录罗马 / 正文阿拉伯）")
        else:
            s1_hdrs = sect_prs[0].findall(f"./{W}headerReference")
            if s1_hdrs:
                s1_t = []
                for hr in s1_hdrs:
                    tgt = rels_map.get(hr.get(f"{R}id", ""), "")
                    if tgt and f"word/{tgt}" in parsed:
                        s1_t.append("".join(t.text or "" for t in parsed[f"word/{tgt}"].iter(f"{W}t")).strip())
                if any(s1_t):
                    sec_issues.append("Section 1（封面/声明）不应显示页眉文字")
            if numbering[0][4]:
                sec_issues.append('封面/声明节的页脚包含 PAGE 页码域')
            if numbering[1][:2] != ('upperRoman', '1') or not numbering[1][4]:
                sec_issues.append('摘要/目录节应从大写罗马 I 起编，并使用 PAGE 页码域')
            if numbering[2][:2] != ('decimal', '1') or not numbering[2][4]:
                sec_issues.append("正文起始/主体节缺少显式 pgNumType/@w:start='1' 的阿拉伯数字页码")
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
        for p in body_paras:
            ps = _p_style(p)
            txt = _p_text(p).strip()
            if not txt:
                continue
            if ps in (CH1_STYLES | H2_STYLES | H3_STYLES | H4_STYLES):
                m_h = re.match(r"^(第[一二三四五六七八九十\d]+章|\d+(?:\.\d+){1,3})([\s\u3000]*)(.+)$", txt)
                if m_h and m_h.group(2) != " ":
                    heading_space_issues.append(f"标题序号空格非单半角: {txt[:35]!r}")
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
                for r in _extract_runs_skip_toc_fields(p):
                    rt = _r_text(r)
                    if not rt or rt.isspace() or "【" in rt:
                        continue
                    if re.search(r"[A-Za-z0-9α-ωΑ-Ωμµ°]", rt):
                        af = styles.effective_ascii_font(r, p).lower()
                        if af not in latin_fonts:
                            font_issues.append(f"英文字体不在当前允许字体中 ({af}): {rt[:30]!r} @ {ps}")

        results.append(CheckResult("Err-9", "常见错误-9(标题序号与标题间单半角空格)", "PASS" if not heading_space_issues else "FAIL", "各级标题序号后空格符合单半角空格要求" if not heading_space_issues else f"发现 {len(heading_space_issues)} 处标题空格异常", heading_space_issues[:10]))
        results.append(CheckResult("Layout-9", "排版-9(正文/摘要/图注/参考文献两端对齐)", "PASS" if not align_issues else "FAIL", "全部核查段落均设为两端对齐" if not align_issues else f"发现 {len(align_issues)} 处未两端对齐", align_issues[:10]))
        results.append(CheckResult("Err-8", "常见错误-8(英文数字与符号用英文字体)", "PASS" if not font_issues else "FAIL", "未发现英文/数字/希腊字母/°C 落入中文字体" if not font_issues else f"发现 {len(font_issues)} 处英文字体偏差", font_issues[:12]))
        results.append(CheckResult('Layout-BodySpacing', '正文固定行距',
                                   'FAIL' if spacing_issues else 'PASS' if body_checked else 'INFO',
                                   f'正文命中已知样式 {body_checked} 段，固定行距偏差 {len(spacing_issues)} 段', spacing_issues[:12]))

        # Locate chapter 1 -> references -> back matter boundaries
        ch1_idx = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and re.match(r"^(?:第一章|Chapter\s+1)\b", _p_text(el).strip(), re.I)), None)
        ref_idx = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and re.fullmatch(r"参考\s*文献|References|Bibliography|Works Cited", _p_text(el).strip(), re.I)), None)
        cv_idx = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and _p_text(el).strip().startswith(("个人简历", "附 录", "附录", "致 谢", "致谢"))), None)
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
            mf = re.match(r"^(?:图|补充图)\s*([1-9A-D])\.(\d+)\s+(.+)$", txt)
            if mf and (ps in FIG_CAP_STYLES or re.match(r"^(?:图|补充图)\s*[1-9A-D]\.\d+\s+\S+", txt)) and len(txt) < 1200:
                ch, num = mf.group(1), int(mf.group(2))
                fig_captions.append((ch, num, p))
                runs = p.findall(f"./{W}r")
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

            mt = re.match(r"^(?:表|附表)\s*([1-9A-D])\.(\d+)\s+(.+)$", txt)
            is_tbl_title = (ps in TBL_TITLE_STYLES) or (mt is not None and not re.search(r"[。；，]\s*$", txt) and "。" not in txt[:35] and len(txt) < 120)
            if mt and is_tbl_title:
                ch, num = mt.group(1), int(mt.group(2))
                tbl_captions.append((ch, num, p))
                runs = p.findall(f"./{W}r")
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
            for m in re.finditer(r"(?:图|补充图)\s*([1-9A-D])\.(\d+)", txt):
                fig_cite_set.add((m.group(1), int(m.group(2))))
            for m in re.finditer(r"(?:表|附表)\s*([1-9A-D])\.(\d+)", txt):
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
            placeholder_labels = ('此处插入矢量科研图', '在此放置作品、资料或研究示意图。')
            if (len(cells) == 1 and top_b is not None and top_b.get(f"{W}val") == 'dashed'
                    and _p_text(tbl).lstrip('【[ ').startswith(placeholder_labels)):
                continue
            content_tbl_count += 1
            issues, needs_word = _table_problems(tbl)
            (unresolved_tables if needs_word else bad_tables).extend(f'表 #{tidx}: {item}' for item in issues)

        results.append(CheckResult("Fig-1", "图表要求-1(按章连续编号)", "PASS" if not seq_issues else "FAIL", f"图={len(fig_captions)}, 表={len(tbl_captions)}", seq_issues))
        results.append(CheckResult("Fig-3-8", "图表要求-3/8(正文图表双向闭环)", "PASS" if not closure_issues else "FAIL", "全部图表与正文双向引用闭环" if not closure_issues else "存在未引用或悬空图表编号", closure_issues))
        results.append(CheckResult("Fig-4", "图表要求-4(图注/表题样式与局部加粗)", "PASS" if not caption_bold_issues else "FAIL", "图注与表题局部加粗规范" if not caption_bold_issues else f"发现 {len(caption_bold_issues)} 处标题加粗问题", caption_bold_issues[:10]))
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
                for ri, r in enumerate(runs):
                    if _is_super_run(r):
                        rt = (_r_text(r) or "").strip()
                        nums = _expand_num_spec(rt)
                        if nums and not re.match(r"^[\d,，\-–—~\s]+$", rt) and not rt.startswith("["):
                            nums = []
                        if nums:
                            cite_groups += 1
                            cited_nums.update(nums)
                            prev_t = _r_text(runs[ri - 1]) if ri > 0 else ""
                            if prev_t and prev_t[-1] in "。，；！？.,;!?" and ps not in TBL_NOTE_STYLES:
                                bad_cite_place.append(f"标点后出现引用上标 {rt!r}: {prev_t[-15:]!r}")
                if re.search(r"\[(\d+(?:[,\-–—]\d+)*)\]", txt) and "【" not in txt:
                    bad_cite_place.append(f"Nature 模式出现方括号引用而非上标: {txt[:40]!r}")
            else:  # gb7714
                for m in re.finditer(r"\[(\d+(?:[,，\-–—]\d+)*)\]", txt):
                    nums = _expand_num_spec(m.group(1))
                    if nums:
                        cite_groups += 1
                        cited_nums.update(nums)

        results.append(
            CheckResult(
                "Ref-1",
                f"引用规范-1({'Nature上标置标点前' if citation_style == 'nature' else 'GB/T 7714方括号引用'})",
                "PASS" if not bad_cite_place else "FAIL",
                f"正文识别 {cite_groups} 处引用组 ({citation_style})" if not bad_cite_place else f"发现 {len(bad_cite_place)} 处引文格式/位置问题",
                bad_cite_place[:10],
            )
        )

        # Bibliography block
        ref_paras: list[ET.Element] = []
        if ref_idx is not None:
            end_r = next((i for i in range(ref_idx + 1, len(blocks)) if blocks[i].tag == f"{W}p" and (_p_style(blocks[i]) in CH1_STYLES or re.match(r"^(个人简历|附\s*录|致\s*谢)", _p_text(blocks[i]).strip()))), len(blocks))
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
            norm_key = re.sub(r"[^a-z0-9\u4e00-\u9fff]", "", rbody.lower())[:40]
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
                if not is_cjk_ref:
                    has_italic = any(styles.effective_italic(r, rp) for r in rp.findall(f"./{W}r") if (_r_text(r) or "").strip())
                    has_bold = any(styles.effective_bold(r, rp) for r in rp.findall(f"./{W}r") if (_r_text(r) or "").strip())
                    if not has_italic or not has_bold:
                        ref_fmt_issues.append(f"文献 {rnum} Nature 格式缺少刊名斜体({has_italic})或卷号加粗({has_bold})")
                    runs = rp.findall(f'./{W}r')
                    ref_fmt_issues.extend(f'文献 {rnum}: {issue}' for issue in _nature_author_title(runs, rp, styles))
                    if re.search(r"\b\d+-\d+\s*\(\d{4}\)", rbody):
                        ref_fmt_issues.append(f"文献 {rnum} 页码使用了短连字符 - 而不是 en-dash –")
                    if re.search(r"\(\d{4}\)\.$", rbody):
                        ref_fmt_issues.append(f"文献 {rnum} 年份括号后不应有多余句点")
                if "doi:" in rbody.lower() or "doi.org/" in rbody.lower():
                    if re.search(r"\(\d{4}\)", rbody) and re.search(r"\d+\s*[–-]\s*\d+", rbody):
                        ref_fmt_issues.append(f"文献 {rnum} 已正式见刊（含卷期页码），请移除冗余 DOI")
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

        results.append(CheckResult("Ref-2-7", f"引用规范-2..7({citation_style}格式/去DOI/序/去重)", "PASS" if not (ref_seq_issues or ref_fmt_issues or ref_lang_order_issues or dup_refs) else "FAIL", f"共核验 {len(ref_nums)} 条文献 (EN.CITE={en_cite_count})", (ref_seq_issues + ref_lang_order_issues + dup_refs + ref_fmt_issues)[:15]))
        results.append(CheckResult("Ref-8", "引用规范-8(正文与参考文献双向闭环)", "PASS" if not ref_closure_issues else "FAIL", "全部文献引用双向闭环" if not ref_closure_issues else "存在未引或越界文献", ref_closure_issues))
        results.append(CheckResult("Ref-9", "引用规范-9(参考文献固定16磅两端对齐)", "PASS" if not ref_style_issues else "FAIL", "全部文献段落符合固定 16 磅与两端对齐" if not ref_style_issues else f"发现 {len(ref_style_issues)} 处行距/对齐偏差", ref_style_issues[:10]))

        # 6. First-person pronouns & consistent self-term (Acad-14)
        first_person_hits: list[str] = []
        self_terms: dict[str, int] = {}
        legal_skip = ("本人郑重声明", "本人声明", "本人完全了解", "本人承诺", "未经作者授权", "版权与使用声明", "论文作者签名")
        for p in all_paras:
            ps = _p_style(p)
            t = _p_text(p).strip()
            if not t or ps in BIB_STYLES or any(k in t for k in legal_skip):
                continue
            if re.search(r"我们", t):
                first_person_hits.append(f"我们@{t[:36]}")
            for term in ("本论文", "本研究", "本工作", "本课题"):
                cnt = len(re.findall(term, t))
                if cnt:
                    self_terms[term] = self_terms.get(term, 0) + cnt
        self_status = "PASS" if (not first_person_hits and len(self_terms) <= 1) else "FAIL"
        results.append(CheckResult("Acad-14", "学术规范-14(禁第一人称/自称一致)", self_status, f"第一人称={len(first_person_hits)}, 自称分布={self_terms}", first_person_hits[:8]))

        # Units, subscripts, Chapter 2 time units, punctuation, enzymes, genes/species, abbreviations
        unit_issues: list[str] = []
        punct_issues: list[str] = []
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
            for formula in ("H2O", "CO2", "Ca2+", "Mg2+", "O2", "N2"):
                if formula in t:
                    unit_issues.append(f"化学式 {formula} 疑未使用上下标格式")
            if ps in (BODY_STYLES | ABS_ZH_STYLES | FIG_CAP_STYLES):
                for m in SUBPANEL_DUNHAO_RE.finditer(t):
                    punct_issues.append(f"多小图图序误用顿号(应为中文逗号): {m.group(0)}")
                for m in CJK_ASCII_PUNCT_RE.finditer(t):
                    punct_issues.append(f"中文行文混用英文标点: {m.group(0)!r}")
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

        results.append(CheckResult("Err-4-5-6", "常见错误-4/5/6(化学式下标/µL/数字单位空格)", "PASS" if not unit_issues else "FAIL", "单位、微升符号及化学式规范" if not unit_issues else f"发现 {len(unit_issues)} 处单位或下标问题", unit_issues[:12]))

        ch2_start = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and re.match(r"^第二章\b", _p_text(el).strip())), None)
        ch3_start = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and re.match(r"^第三章\b", _p_text(el).strip())), None)
        has_cjk_t, has_eng_t = False, False
        if ch2_start is not None:
            ch2_txt = "\n".join(_p_text(el) for el in blocks[ch2_start : (ch3_start or body_end)])
            has_cjk_t = bool(CJK_TIME_RE.search(ch2_txt))
            has_eng_t = bool(ENG_TIME_RE.search(ch2_txt))
        results.append(CheckResult("Err-7", "常见错误-7(第二章时间单位中英文统一)", "FAIL" if (has_cjk_t and has_eng_t) else "PASS", f"第二章中文时间={has_cjk_t}, 英文时间={has_eng_t}", ["第二章同时混用中文时间词与英文缩写"] if (has_cjk_t and has_eng_t) else []))
        results.append(CheckResult("Err-2-11-12", "常见错误-2/11/12(中英文标点与小图逗号)", "PASS" if not punct_issues else "FAIL", "中英文标点及多小图逗号规范" if not punct_issues else f"发现 {len(punct_issues)} 处标点混用", punct_issues[:12]))
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
            ack_i = next((i for i, el in enumerate(blocks) if el.tag == f"{W}p" and re.fullmatch(r"致\s*谢", _p_text(el).strip())), None)
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
            and not re.match(r"^(?:图|补充图|表|附表)\s*[1-9A-D]", _p_text(p).strip())
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
        if isinstance(wc_range, list) and len(wc_range) == 2:
            lo, hi = wc_range[0], wc_range[1]
            if (lo is not None and total_w < lo) or (hi is not None and total_w > hi):
                wc_status = "WARN"
                bound_str = f"{lo:,}–{hi:,}" if (lo is not None and hi is not None) else f">= {lo:,}"
                wc_msg += f" — 当前模式建议 {bound_str} 字"
        results.append(CheckResult("WordCount", "正文字数统计", wc_status, wc_msg, []))

        custom_style = False
        if args.profile_config:
            override = json.loads(pathlib.Path(args.profile_config).read_text(encoding='utf-8'))
            custom_style = override.get('citation_style') in ('nature', 'gb7714')
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
        ("rendered-figures-tables", "如有图表，在最终 PDF 中逐一检查图注与图片是否同页、表格跨页表头是否重复、表线和文字是否清晰。"),
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


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only OOXML & 65-item compliance audit for theses and reports.")
    add_arguments(parser)
    args = parser.parse_args(argv)
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


if __name__ == "__main__":
    raise SystemExit(main())
