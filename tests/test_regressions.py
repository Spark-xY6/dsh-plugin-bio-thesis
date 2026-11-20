"""Regression tests for external templates and known false passes. Uses synthetic Word files only."""

from pathlib import Path
import json
import subprocess
import sys
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent / 'skills/bio-thesis-formatter'
CLI = SKILL / 'scripts/bio_thesis.py'
MASTER = SKILL / 'resources/bio-thesis-master.docx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
DC = '{http://purl.org/dc/elements/1.1/}'
MAIN = 'word/document.xml'


def visible(node):
    return ''.join(t.text or '' for t in node.iter(W + 't'))


def modify(source, target, changes):
    with ZipFile(source) as src, ZipFile(target, 'w', ZIP_DEFLATED) as dst:
        for name in src.namelist():
            data = src.read(name)
            if name in changes:
                root = ET.fromstring(data)
                changes[name](root)
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            dst.writestr(name, data)


def change_numbering(root):
    cover, front, _ = root.iter(W + 'sectPr')
    source = {ref.get(W + 'type'): ref.get(R + 'id') for ref in front.findall(W + 'footerReference')}
    for ref in cover.findall(W + 'footerReference'):
        ref.set(R + 'id', source[ref.get(W + 'type')])
    front.find(W + 'pgNumType').set(W + 'start', '7')


def change_borders(root, remove=False):
    for table in root.iter(W + 'tbl'):
        props = table.find(W + 'tblPr')
        border = props.find(W + 'tblBorders') if props is not None else None
        if border is None or border.find(W + 'top') is None or border.find(W + 'top').get(W + 'val') != 'single':
            continue
        if remove:
            props.remove(border)
        else:
            border.find(W + 'top').set(W + 'sz', '2')
            border.find(W + 'insideH').set(W + 'val', 'single')
        return
    raise AssertionError('No table in synthetic fixture')


def drop_bibliography(root):
    body = root.find(W + 'body')
    blocks = list(body)
    start = next(i for i, item in enumerate(blocks) if visible(item) == '参考文献')
    end = next(i for i, item in enumerate(blocks) if visible(item).startswith('个人简历'))
    for element in blocks[start:end]:
        body.remove(element)


def change_line(root):
    for style in root.iter(W + 'style'):
        if style.get(W + 'styleId') == 'BodyTextZH':
            style.find(f'./{W}pPr/{W}spacing').set(W + 'line', '240')
            return
    raise AssertionError('No body style in synthetic fixture')


def change_nature(root):
    text = next(t for t in root.iter(W + 't') if 'Example, A. B.' in (t.text or ''))
    text.text = text.text.replace('Example, A. B.', 'Example, AB.').replace(' & ', ' and ')
    title = 'A fictional example for document layout.'
    text.text = text.text.replace(title, title.upper())


def custom_body(root):
    body = root.find(W + 'body')
    for child in list(body)[:-1]:
        body.remove(child)
    section = body.find(W + 'sectPr')
    section.find(W + 'pgMar').set(W + 'top', '2281')
    p = ET.Element(W + 'p')
    for piece in ('Prefix {{title_', 'zh}} and {{title_zh}} / {{author}}'):
        ET.SubElement(ET.SubElement(p, W + 'r'), W + 't').text = piece
    body.insert(0, p)


def use_template_type(root):
    for child in root:
        if child.get('PartName') == '/word/document.xml':
            child.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml')


class RegressionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.work = Path(self.tmp.name)

    def invoke(self, *args, returncode=0):
        result = subprocess.run([sys.executable, '-B', '-X', 'utf8', str(CLI), *map(str, args)],
                                capture_output=True, encoding='utf-8', check=False)
        self.assertEqual(result.returncode, returncode, result.stdout + '\n' + result.stderr)
        return result

    def audit(self, doc, *args, returncode=0):
        report = self.work / 'report.json'
        self.invoke('audit', '--docx', doc, '--json-report', report, *args, returncode=returncode)
        findings = json.loads(report.read_text(encoding='utf-8'))['checks']
        return {c['check_id']: c for c in findings}

    def test_known_false_passes(self):
        defects = {
            'numbering': (MAIN, change_numbering, 'Layout-SecNum'),
            'borders': (MAIN, change_borders, 'Tbl-3Line'),
            'missing-borders': (MAIN, lambda r: change_borders(r, True), 'Tbl-3Line'),
            'spacing': ('word/styles.xml', change_line, 'Layout-BodySpacing'),
            'bibliography': (MAIN, drop_bibliography, 'Ref-8'),
            'nature': (MAIN, change_nature, 'Ref-2-7'),
            'creator': ('docProps/core.xml', lambda r: setattr(r.find(DC + 'creator'), 'text', 'DummyAuditIdentity'), 'Acad-10-Blind'),
        }
        valid = self.work / 'blind.docx'
        self.invoke('init', '--output', valid, '--review-stage', 'blind', '--preset', 'cn-lifesci-default')
        opts = ('--review-stage', 'blind', '--preset', 'cn-lifesci-default', '--sensitive-names', 'dummyauditidentity')
        self.audit(valid, *opts)
        for label, (part, edit, check_id) in defects.items():
            with self.subTest(label=label):
                sample = self.work / (label + '.docx')
                modify(valid, sample, {part: edit})
                self.assertEqual(self.audit(sample, *opts, returncode=1)[check_id]['status'], 'FAIL')

    def test_external_docx_and_dotx_keep_layout_and_non_target_parts(self):
        for suffix in ('docx', 'dotx'):
            with self.subTest(suffix=suffix):
                source = self.work / f'custom.{suffix}'
                output = self.work / f'candidate-{suffix}.docx'
                changes = {MAIN: custom_body}
                if suffix == 'dotx':
                    changes['[Content_Types].xml'] = use_template_type
                modify(MASTER, source, changes)
                before = source.read_bytes()
                self.invoke('init', '--template', source, '--output', output, '--title-zh', '合成 <&> 标题',
                            '--author', 'Dummy Person', '--review-stage', 'blind')
                self.assertEqual(before, source.read_bytes())
                with ZipFile(source) as src, ZipFile(output) as dst:
                    self.assertEqual(src.namelist(), dst.namelist())
                    allowed = {MAIN, '[Content_Types].xml'} if suffix == 'dotx' else {MAIN}
                    self.assertEqual({m for m in src.namelist() if src.read(m) != dst.read(m)}, allowed)
                    self.assertIn('Prefix 合成 <&> 标题 and 合成 <&> 标题 / 【匿名评审隐去】',
                                  visible(ET.fromstring(dst.read(MAIN))))
                result = self.audit(output, '--template', source, '--review-stage', 'blind',
                                    '--sensitive-names', 'Dummy Person')
                for check_id in ('Layout-1', 'Layout-SecNum', 'Layout-TemplateStyles'):
                    self.assertEqual(result[check_id]['status'], 'PASS')
                self.assertEqual(result['Layout-TOC']['status'], 'INFO')
                self.assertEqual(result['Tbl-3Line']['status'], 'INFO')
                self.invoke('init', '--template', output, '--output', output, '--overwrite', returncode=1)

    def test_output_aliases_and_consistent_profile(self):
        profile = self.work / 'profile.json'
        profile.write_text(json.dumps({'page': {'margins_cm': {'top': 4.0}},
                                       'typography': {'body_line_pt': 28, 'cjk_body_font': '楷体'}}), encoding='utf-8')
        doc = self.work / 'custom.docx'
        self.invoke('init', '--output', doc, '--profile-config', profile)
        result = self.audit(doc, '--profile-config', profile)
        self.assertEqual(result['Layout-1']['status'], 'PASS')
        self.assertEqual(result['Layout-BodySpacing']['status'], 'PASS')
        original = doc.read_bytes()
        before_profile = profile.read_bytes()
        self.invoke('audit', '--docx', doc, '--json-report', doc, returncode=2)
        self.invoke('audit', '--docx', doc, '--base-docx', doc, '--json-report', doc, returncode=2)
        self.invoke('audit', '--docx', doc, '--profile-config', profile, '--json-report', profile, returncode=2)
        hardlink = self.work / 'hard-link.json'
        hardlink.hardlink_to(doc)
        self.invoke('audit', '--docx', doc, '--json-report', hardlink, returncode=2)
        self.assertEqual(doc.read_bytes(), original)
        self.assertEqual(profile.read_bytes(), before_profile)

    def test_gb_journal_metadata(self):
        reference = self.work / 'gb.docx'
        self.invoke('init', '--output', reference, '--preset', 'cn-gb7714-general', '--mode', 'proposal')
        candidate = self.work / 'missing-metadata.docx'

        def strip_metadata(root):
            node = next(t for t in root.iter(W + 't') if (t.text or '').startswith('[1] EXAMPLE'))
            node.text = '[1] Unspecified entry[J].'

        modify(reference, candidate, {MAIN: strip_metadata})
        checks = self.audit(candidate, '--citation-style', 'gb7714', '--mode', 'proposal', returncode=1)
        self.assertEqual(checks['Ref-2-7']['status'], 'FAIL')


if __name__ == '__main__':
    unittest.main()
