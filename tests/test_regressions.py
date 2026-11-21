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

    def test_direct_format_policy_scope_and_locations(self):
        # Independent style IDs for a synthetic field-report template.
        def field_styles(root):
            root.clear()
            for sid, kind in [('FieldNotes', 'paragraph'), ('FieldDetail', 'paragraph'),
                              ('MapLegend', 'paragraph'), ('KeyPhrase', 'character')]:
                style = ET.SubElement(root, W + 'style', {W + 'styleId': sid, W + 'type': kind})
                if sid == 'FieldNotes':
                    style.set(W + 'default', '1')
                if sid == 'FieldDetail':
                    ET.SubElement(style, W + 'basedOn', {W + 'val': 'FieldNotes'})

        def field_body(root):
            body = root.find(W + 'body')
            section = body.find(W + 'sectPr')
            body.clear()
            for xml in (
                '<p><pPr><rPr><sz val="36"/></rPr></pPr>'
                '<r><rPr><rFonts hint="eastAsia"/><b/><i/></rPr><t>Survey</t></r></p>',
                '<p><pPr><pStyle val="MapLegend"/></pPr><r><rPr><sz val="40"/></rPr><t>Legend</t></r></p>',
                '<tbl><tr><tc><p><pPr><pStyle val="FieldDetail"/></pPr>'
                '<r><rPr><rStyle val="KeyPhrase"/></rPr><t>Field reading</t></r></p></tc></tr></tbl>',
            ):
                # Attribute names in WordprocessingML are also namespace-qualified.
                block = ET.fromstring(xml)
                for node in block.iter():
                    node.tag = W + node.tag
                    for key in list(node.attrib):
                        node.set(W + key, node.attrib.pop(key))
                body.append(block)
            body.append(section)

        source = self.work / 'field-template.docx'
        modify(MASTER, source, {'word/styles.xml': field_styles, MAIN: field_body})
        config = self.work / 'rules.json'
        config.write_text(json.dumps({'audit': {'direct_formatting': [
            {'id': 'field-font', 'paragraph_style': 'FieldNotes', 'forbid': ['font_size', 'font_face', 'color']},
        ]}}), encoding='utf-8')
        options = ('--template', source, '--profile-config', config)
        check_id = 'DirectFormat-field-font'
        self.assertEqual(self.audit(source, *options)[check_id]['status'], 'PASS')

        def introduce_overrides(root):
            first, _, third = root.iter(W + 'p')
            ET.SubElement(first.find(f'./{W}r/{W}rPr'), W + 'sz', {W + 'val': '21'})
            props = third.find(f'./{W}r/{W}rPr')
            for tag, attr, value in [('szCs', 'val', '31'), ('rFonts', 'eastAsia', 'Example Font'),
                                     ('color', 'val', 'AA1122')]:
                ET.SubElement(props, W + tag, {W + attr: value})

        changed = self.work / 'field-changes.docx'
        modify(source, changed, {MAIN: introduce_overrides})
        self.assertNotIn(check_id, self.audit(changed, '--template', source))
        finding = self.audit(changed, *options, returncode=1)[check_id]
        self.assertEqual(finding['status'], 'FAIL')
        self.assertEqual(finding['details'], ['paragraph=1 properties=font_size',
                         'paragraph=3 properties=font_size,font_face,color'])

        def insert_history(root):
            para = next(root.iter(W + 'p'))
            for change in ('del', 'moveFrom'):
                run = ET.SubElement(ET.SubElement(para, W + change), W + 'r')
                ET.SubElement(ET.SubElement(run, W + 'rPr'), W + 'sz', {W + 'val': '11'})
                ET.SubElement(run, W + 'delText').text = 'Old reading'
            props = para.find(f'./{W}r/{W}rPr')
            old = ET.SubElement(ET.SubElement(props, W + 'rPrChange'), W + 'rPr')
            ET.SubElement(old, W + 'color', {W + 'val': '000000'})
            table = next(root.iter(W + 'tbl'))
            deleted_row = ET.SubElement(table, W + 'tr')
            ET.SubElement(ET.SubElement(deleted_row, W + 'trPr'), W + 'del')
            old_para = ET.SubElement(ET.SubElement(deleted_row, W + 'tc'), W + 'p')
            old_run = ET.SubElement(old_para, W + 'r')
            ET.SubElement(ET.SubElement(old_run, W + 'rPr'), W + 'sz', {W + 'val': '11'})
            ET.SubElement(old_run, W + 't').text = 'Retired record'

        revised = self.work / 'field-history.docx'
        modify(source, revised, {MAIN: insert_history})
        findings = self.audit(revised, *options, returncode=1)
        self.assertEqual(findings['Pkg-Clean']['status'], 'FAIL')
        self.assertEqual(findings[check_id]['status'], 'PASS')

    def test_direct_format_policy_rejects_mistyped_rules(self):
        config = self.work / 'rules.json'
        rule = {'id': 'body', 'paragraph_style': 'BodyTextZH', 'forbid': ['font_size']}
        for key, rules, message in (
                ('direct_formating', [rule], 'Unknown audit option'),
                ('direct_formatting', [{**rule, 'forbid': ['szze']}], '.forbid'),
                ('direct_formatting', [{**rule, 'paragraph_style': 'AbsentStyle'}], '.paragraph_style'),
                ('direct_formatting', [{**rule, 'typo': True}], 'requires exactly'),
                ('direct_formatting', [rule, rule], '.id must be unique')):
            with self.subTest(key=key, rules=rules):
                config.write_text(json.dumps({'audit': {key: rules}}), encoding='utf-8')
                failed = self.invoke('audit', '--docx', MASTER, '--profile-config', config, returncode=2)
                self.assertIn(message, failed.stderr)

    def test_existing_live_citation_switch(self):
        source = self.work / 'draft.docx'
        self.invoke('init', '--output', source)
        doc = self.work / 'live-field.docx'

        def add_field(root):
            paragraph = next(root.iter(W + 'p'))
            ET.SubElement(paragraph, W + 'fldSimple', {W + 'instr': 'ADDIN EN.CITE'})

        modify(source, doc, {MAIN: add_field})
        config = self.work / 'rules.json'
        for allow in (True, False):
            config.write_text(json.dumps({'audit': {'allow_live_citations': allow}}), encoding='utf-8')
            findings = self.audit(doc, '--profile-config', config, returncode=0 if allow else 1)
            self.assertEqual(findings['Ref-2-7']['status'], 'PASS' if allow else 'FAIL')

    def test_required_package_parts_and_absolute_internal_targets(self):
        source = self.work / 'package-source.docx'
        self.invoke('init', '--output', source)
        for part, finding in [('[Content_Types].xml', 'Pkg-XML'), ('_rels/.rels', 'Pkg-Rels')]:
            with self.subTest(part=part):
                broken = self.work / 'missing-part.docx'
                with ZipFile(source) as src, ZipFile(broken, 'w') as dst:
                    for name in src.namelist():
                        if name != part:
                            dst.writestr(name, src.read(name))
                self.assertEqual(self.audit(broken, returncode=1)[finding]['status'], 'FAIL')
                rejected = self.invoke('init', '--template', broken, '--output', self.work / 'bad.docx', returncode=1)
                self.assertIn(part, rejected.stderr)
                self.assertFalse((self.work / 'bad.docx').exists())
        wrong = self.work / 'wrong-main.docx'
        def misdirect(root):
            rel = next(r for r in root if r.get('Type', '').endswith('/officeDocument'))
            rel.set('Target', 'word/styles.xml')
        modify(source, wrong, {'_rels/.rels': misdirect})
        self.assertEqual(self.audit(wrong, returncode=1)['Pkg-Rels']['status'], 'FAIL')

        def absolute_targets(root):
            for rel in root:
                if rel.get('TargetMode') != 'External' and '/' not in rel.get('Target', ''):
                    rel.set('Target', '/word/' + rel.get('Target'))
        absolute = self.work / 'absolute-rels.docx'
        modify(source, absolute, {'word/_rels/document.xml.rels': absolute_targets})
        checks = self.audit(absolute, '--template', source)
        self.assertEqual(checks['Pkg-Rels']['status'], 'PASS')
        self.assertEqual(checks['Layout-SecNum']['status'], 'PASS')

    def test_chemical_subscripts_and_ambiguous_plain_text(self):
        source = self.work / 'chem-template.docx'
        modify(MASTER, source, {MAIN: custom_body})
        for formatting, expected in [(True, 'PASS'), (False, 'WARN')]:
            def insert_formula(root):
                paragraph = next(root.iter(W + 'p'))
                for value, vertical in [(' H', None), ('2', 'subscript'), ('O / Mg', None),
                                        ('2+', 'superscript'), (' / CO2sensor', None)]:
                    run = ET.SubElement(paragraph, W + 'r')
                    if formatting and vertical:
                        ET.SubElement(ET.SubElement(run, W + 'rPr'), W + 'vertAlign', {W + 'val': vertical})
                    ET.SubElement(run, W + 't').text = value
            candidate = self.work / 'formula.docx'
            modify(source, candidate, {MAIN: insert_formula})
            finding = self.audit(candidate, '--template', source)['Err-4-5-6']
            self.assertEqual(finding['status'], expected)
            self.assertEqual(len(finding['details']), 0 if formatting else 2)


if __name__ == '__main__':
    unittest.main()
