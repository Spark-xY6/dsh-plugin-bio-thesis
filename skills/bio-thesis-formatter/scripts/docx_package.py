"""Word package I/O and explicit template fields for the thesis formatter."""

from __future__ import annotations

from collections import Counter
from contextlib import contextmanager
import os
from pathlib import Path
import posixpath
import re
import tempfile
import unicodedata
from xml.dom import Node, minidom
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
DOTX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
STORY = re.compile(r'word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$')


def same_path(a: Path, b: Path) -> bool:
    return a.resolve() == b.resolve() or (a.exists() and b.exists() and a.samefile(b))


@contextmanager
def staged_file(destination: Path, overwrite: bool = True):
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=destination.parent, delete=False) as temporary:
        stage = Path(temporary.name)
    try:
        yield stage
        if overwrite:
            os.replace(stage, destination)
        elif os.name == 'nt':
            # Windows rename is atomic and refuses to replace an existing destination.
            os.rename(stage, destination)
        else:
            # Atomic creation that fails if another process already published this name.
            os.link(stage, destination)
    finally:
        stage.unlink(missing_ok=True)


def write_package(destination: Path, items: dict[str, bytes], overwrite: bool) -> None:
    with staged_file(destination, overwrite) as stage:
        with ZipFile(stage, 'w', ZIP_DEFLATED) as package:
            for name, data in items.items():
                package.writestr(name, data)


def resolve_target(source_part: str, target: str) -> str:
    """Resolve an internal relationship to a ZIP member relative to its source part."""
    path = target.partition('#')[0]
    if not path:
        return source_part if target.startswith('#') else ''
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), path)).lstrip('/')


def word_main_type(types: ET.Element | None) -> str:
    if types is None or types.tag != f'{{{CT}}}Types':
        raise ValueError('Missing or invalid [Content_Types].xml')
    mains = [el.get('ContentType') for el in types.findall(f'{{{CT}}}Override')
             if el.get('PartName') == '/word/document.xml']
    if not mains:
        mains = [el.get('ContentType') for el in types.findall(f'{{{CT}}}Default')
                 if el.get('Extension', '').lower() == 'xml']
    if len(mains) != 1 or mains[0] not in (DOCX, DOTX):
        raise ValueError('[Content_Types].xml must declare a supported Word main part')
    return mains[0]


def has_main_relationship(root: ET.Element | None) -> bool:
    if root is None or root.tag != f'{{{PACKAGE_RELS}}}Relationships':
        return False
    mains = [r for r in root.findall(f'{{{PACKAGE_RELS}}}Relationship') if r.get('Type') == OFFICE_DOCUMENT]
    return (len(mains) == 1 and bool(mains[0].get('Id'))
            and mains[0].get('TargetMode', 'Internal') == 'Internal'
            and resolve_target('', mains[0].get('Target', '')) == 'word/document.xml')


def read_template(path: Path) -> tuple[dict[str, bytes], str]:
    if path.suffix.lower() not in {'.docx', '.dotx'}:
        raise ValueError('Template must be a macro-free .docx or .dotx')
    with ZipFile(path) as package:
        names = package.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Template has duplicate ZIP member names')
        items = {name: package.read(name) for name in names}
    if any(name.lower().endswith('vbaproject.bin') for name in items):
        raise ValueError('Macro-enabled templates are unsupported')
    if any(name.lower().startswith('_xmlsignatures/') for name in items):
        raise ValueError('Signed templates are unsupported because edits invalidate the signature')
    main_type = word_main_type(ET.fromstring(items['[Content_Types].xml']) if '[Content_Types].xml' in items else None)
    if not has_main_relationship(ET.fromstring(items['_rels/.rels']) if '_rels/.rels' in items else None):
        raise ValueError('_rels/.rels must contain one internal officeDocument relationship to word/document.xml')
    if (main_type == DOTX) != (path.suffix.lower() == '.dotx'):
        raise ValueError('Template suffix and Word package type disagree')
    root = ET.fromstring(items['word/document.xml'])
    if root.tag != W + 'document' or root.find(W + 'body') is None:
        raise ValueError('Template has no supported Word document body')
    return items, main_type


def _children(node):
    return [child for child in node.childNodes if child.nodeType == Node.ELEMENT_NODE]


def _group_text(node) -> str:
    return ''.join(child.data for child in node.childNodes if child.nodeType == Node.TEXT_NODE)


def _replace_group(nodes, token_pattern, values, filled):
    originals = [_group_text(node) for node in nodes]
    combined = ''.join(originals)
    spans = list(token_pattern.finditer(combined))
    if not spans:
        return
    offsets = []
    running = 0
    for value in originals:
        offsets.append(running)
        running += len(value)
    texts = originals.copy()
    for match in reversed(spans):
        name = match.group(1)
        for index, (offset, original) in enumerate(zip(offsets, originals)):
            end = offset + len(original)
            if end <= match.start() or offset >= match.end():
                continue
            left = max(0, match.start() - offset)
            right = min(len(original), match.end() - offset)
            insert = values[name] if offset <= match.start() < end else ''
            texts[index] = texts[index][:left] + insert + texts[index][right:]
        filled[name] += 1
    for node, original, new_text in zip(nodes, originals, texts):
        if new_text == original:
            continue
        for child in list(node.childNodes):
            node.removeChild(child)
        node.appendChild(node.ownerDocument.createTextNode(new_text))
        if new_text != new_text.strip():
            node.setAttributeNS(XML_NS, 'xml:space', 'preserve')


def fill_template(items: dict[str, bytes], package_type: str, values: dict[str, str]) -> Counter:
    if package_type == DOTX:
        types = minidom.parseString(items['[Content_Types].xml'])
        for override in types.getElementsByTagNameNS(CT, 'Override'):
            if override.getAttribute('PartName') == '/word/document.xml':
                override.setAttribute('ContentType', DOCX)
                break
        else:
            prefix = types.documentElement.prefix
            override = types.createElementNS(CT, f'{prefix}:Override' if prefix else 'Override')
            override.setAttribute('PartName', '/word/document.xml')
            override.setAttribute('ContentType', DOCX)
            types.documentElement.appendChild(override)
        items['[Content_Types].xml'] = types.toxml(encoding='utf-8')
    filled: Counter = Counter()
    if not values:
        return filled
    tokens = re.compile(r'\{\{(' + '|'.join(map(re.escape, values)) + r')\}\}')
    for name, data in list(items.items()):
        if not STORY.fullmatch(name):
            continue
        part = minidom.parseString(data)
        before = filled.total()
        field_depth = 0
        for paragraph in part.getElementsByTagNameNS(W_NS, 'p'):
            group = []
            for run in _children(paragraph):
                if run.namespaceURI == W_NS and run.localName == 'pPr':
                    continue
                direct = _children(run)
                simple = run.namespaceURI == W_NS and run.localName == 'r' and all(
                    node.namespaceURI == W_NS and node.localName in ('rPr', 't') for node in direct
                )
                if simple and not field_depth:
                    group.extend(node for node in direct if node.localName == 't')
                    continue
                _replace_group(group, tokens, values, filled)
                group.clear()
                for mark in run.getElementsByTagNameNS(W_NS, 'fldChar'):
                    kind = mark.getAttributeNS(W_NS, 'fldCharType')
                    if kind == 'begin':
                        field_depth += 1
                    elif kind == 'end':
                        field_depth = max(0, field_depth - 1)
            _replace_group(group, tokens, values, filled)
        if filled.total() > before:
            items[name] = part.toxml(encoding='utf-8')
    return filled


def _normalize(value: str) -> str:
    return unicodedata.normalize('NFKC', value).casefold()


def identity_parts(parsed: dict[str, ET.Element], tokens: list[str]) -> list[str]:
    needles = [_normalize(token.strip()) for token in tokens if token.strip()]
    if not needles:
        return []
    hits = []
    visible = {W + tag for tag in ('t', 'delText', 'instrText', 'delInstrText')}
    for name, root in parsed.items():
        fields = [value for el in root.iter() for value in (*el.attrib.values(), el.text or '')]
        fields.extend(''.join(node.text or '' for node in paragraph.iter() if node.tag in visible)
                      for paragraph in root.iter(W + 'p'))
        if any(needle in _normalize(value) for value in fields for needle in needles):
            hits.append(name)
    return hits
