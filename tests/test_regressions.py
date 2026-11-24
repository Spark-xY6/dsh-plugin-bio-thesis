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

    def test_margin_preview_and_copy_preserve_other_content(self):
        sample = self.work / '论文 初稿.docx'
        output = self.work / '修正版.docx'
        template = self.work / '学院模板.dotx'
        modify(MASTER, template, {'[Content_Types].xml': use_template_type})

        def changed(root):
            first, second, third = root.iter(W + 'sectPr')
            second.find(W + 'pgMar').set(W + 'left', '500')
            second.find(W + 'pgMar').set(W + 'bottom', '-600')
            third.find(W + 'pgMar').set(W + 'top', '750')
            # These two properties belong to the manuscript.
            first.find(W + 'pgMar').set(W + 'gutter', '180')
            next(root.iter(W + 't')).text = 'Author prose with & and < preserved.'

        modify(MASTER, sample, {MAIN: changed})
        originals = (sample.read_bytes(), template.read_bytes())
        args = ('repair-margins', '--docx', sample, '--template', template)
        before = set(self.work.iterdir())
        self.assertIn('MARGIN_CHANGES=3', self.invoke(*args).stdout)
        self.assertEqual(before, set(self.work.iterdir()))
        self.assertIn('MARGIN_CHANGES=3', self.invoke(*args, '--output', output).stdout)
        self.assertEqual(originals, (sample.read_bytes(), template.read_bytes()))
        with ZipFile(sample) as old, ZipFile(output) as new, ZipFile(template) as model:
            self.assertEqual(old.namelist(), new.namelist())
            self.assertEqual([n for n in old.namelist() if old.read(n) != new.read(n)], [MAIN])
            original_root = ET.fromstring(old.read(MAIN))
            model_root = ET.fromstring(model.read(MAIN))
            for dest, src in zip(original_root.iter(W + 'pgMar'), model_root.iter(W + 'pgMar')):
                for side in ('left', 'right', 'top', 'bottom'):
                    dest.set(W + side, src.get(W + side))
            self.assertEqual(ET.tostring(original_root), ET.tostring(ET.fromstring(new.read(MAIN))))
        self.assertIn('MARGIN_CHANGES=0', self.invoke('repair-margins', '--docx', output, '--template', template).stdout)
        for collision in (sample, template, output):
            prior = collision.read_bytes()
            self.invoke(*args, '--output', collision, returncode=1)
            self.assertEqual(prior, collision.read_bytes())

    def test_margin_repair_refuses_ambiguous_sections(self):
        def drop_section(root):
            ppr = next(p for p in root.iter(W + 'pPr') if p.find(W + 'sectPr') is not None)
            ppr.remove(ppr.find(W + 'sectPr'))

        def tracked_section(root):
            section = next(root.iter(W + 'sectPr'))
            ET.SubElement(ET.SubElement(section, W + 'sectPrChange'), W + 'sectPr')

        def ambiguous_margin(root):
            next(root.iter(W + 'pgMar')).attrib.pop(W + 'left')

        for edit in (drop_section, tracked_section, ambiguous_margin):
            sample = self.work / (edit.__name__ + '.docx')
            output = self.work / 'rejected.docx'
            modify(MASTER, sample, {MAIN: edit})
            self.invoke('repair-margins', '--docx', sample, '--template', MASTER, '--output', output, returncode=1)
            self.assertFalse(output.exists())

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

    def test_role_aligned_layout_repair_with_landscape_island_and_numbering(self):
        template = self.work / 'school-template.docx'
        self.invoke('init', '--output', template)
        draft = self.work / 'five-section-draft.docx'

        def build_five_section_with_island(root):
            body = root.find(W + 'body')
            sect_list = list(body.iter(W + 'sectPr'))
            cover_sect, front_sect, body_sect = sect_list[0], sect_list[1], sect_list[2]
            for sect in (cover_sect, front_sect, body_sect):
                mar = sect.find(W + 'pgMar')
                mar.set(W + 'top', '1134')
                mar.set(W + 'header', '500')
            pgnum = body_sect.find(W + 'pgNumType')
            if pgnum is None:
                pgnum = ET.SubElement(body_sect, W + 'pgNumType')
            pgnum.set(W + 'start', '4')

            # Split body into portrait body -> landscape table island -> trailing portrait body
            body_idx = list(body).index(body_sect)
            p_island_pre = ET.Element(W + 'p')
            ppr1 = ET.SubElement(p_island_pre, W + 'pPr')
            s1 = ET.SubElement(ppr1, W + 'sectPr')
            ET.SubElement(s1, W + 'pgSz', {W + 'w': '11906', W + 'h': '16838'})
            ET.SubElement(s1, W + 'pgMar', {
                W + 'top': '1134', W + 'right': '1701', W + 'bottom': '1417',
                W + 'left': '1701', W + 'header': '851', W + 'footer': '992', W + 'gutter': '0',
            })
            ET.SubElement(ET.SubElement(p_island_pre, W + 'r'), W + 't').text = '2.1 实验设计'

            wide_tbl = ET.Element(W + 'tbl')
            tr = ET.SubElement(wide_tbl, W + 'tr')
            ET.SubElement(ET.SubElement(ET.SubElement(tr, W + 'tc'), W + 'p'), W + 'r')

            p_island_end = ET.Element(W + 'p')
            ppr2 = ET.SubElement(p_island_end, W + 'pPr')
            s2 = ET.SubElement(ppr2, W + 'sectPr')
            ET.SubElement(s2, W + 'pgSz', {W + 'w': '16838', W + 'h': '11906', W + 'orient': 'landscape'})
            ET.SubElement(s2, W + 'pgMar', {
                W + 'top': '1000', W + 'right': '1000', W + 'bottom': '1000',
                W + 'left': '1000', W + 'header': '500', W + 'footer': '500', W + 'gutter': '0',
            })

            body.insert(body_idx, p_island_pre)
            body.insert(body_idx + 1, wide_tbl)
            body.insert(body_idx + 2, p_island_end)

        def break_mirror(root):
            node = root.find(W + 'mirrorMargins')
            if node is not None:
                root.remove(node)

        modify(template, draft, {MAIN: build_five_section_with_island, 'word/settings.xml': break_mirror})
        before = self.audit(draft, '--template', template, returncode=1)
        self.assertEqual(before['Layout-Settings']['status'], 'FAIL')
        self.assertEqual(before['Layout-1']['status'], 'FAIL')
        self.assertEqual(before['Layout-SecNum']['status'], 'FAIL')
        self.assertNotEqual(before['Fig-1']['status'], 'INFO')
        self.assertNotEqual(before['Fig-3-8']['status'], 'INFO')
        self.assertNotEqual(before['WordCount']['status'], 'FAIL')

        repaired = self.work / 'five-section-repaired.docx'
        self.invoke('repair', '--docx', draft, '--template', template, '--output', repaired)
        after = self.audit(repaired, '--template', template, '--base-docx', draft)
        self.assertEqual(after['Layout-1']['status'], 'PASS')
        self.assertEqual(after['Layout-Settings']['status'], 'PASS')
        self.assertEqual(after['Layout-SecNum']['status'], 'PASS')
        self.assertEqual(after['WordCount']['status'], 'INFO')

        with ZipFile(repaired) as zf:
            rep_root = ET.fromstring(zf.read(MAIN))
        sects = list(rep_root.iter(W + 'sectPr'))
        self.assertEqual(len(sects), 5)
        island_sz = sects[3].find(W + 'pgSz')
        self.assertGreater(int(island_sz.get(W + 'w')), int(island_sz.get(W + 'h')))

    def test_sdt_autonumbered_chapters_theme_fonts_brackets_and_acknowledgements(self):
        template = self.work / 'modern-template.docx'
        self.invoke('init', '--output', template)
        modern = self.work / 'modern-draft.docx'

        def build_modern_elements(root):
            body = root.find(W + 'body')
            children_list = list(body)
            ref_para = next(
                el for el in children_list
                if el.tag == W + 'p' and '参考文献' in ''.join(t.text or '' for t in el.iter(W + 't'))
            )
            insert_at = children_list.index(ref_para)

            # Content control <w:sdt> wrapping a valid three-line table inside the main prose
            sdt = ET.Element(W + 'sdt')
            sdt_c = ET.SubElement(sdt, W + 'sdtContent')
            tbl = ET.Element(W + 'tbl')
            tbl_pr = ET.SubElement(tbl, W + 'tblPr')
            borders = ET.SubElement(tbl_pr, W + 'tblBorders')
            ET.SubElement(borders, W + 'top', {W + 'val': 'single', W + 'sz': '12'})
            ET.SubElement(borders, W + 'bottom', {W + 'val': 'single', W + 'sz': '12'})
            for r_idx in range(2):
                tr = ET.SubElement(tbl, W + 'tr')
                if r_idx == 0:
                    tr_pr = ET.SubElement(tr, W + 'trPr')
                    ET.SubElement(tr_pr, W + 'tblHeader')
                tc = ET.SubElement(tr, W + 'tc')
                tc_pr = ET.SubElement(tc, W + 'tcPr')
                tc_b = ET.SubElement(tc_pr, W + 'tcBorders')
                if r_idx == 0:
                    ET.SubElement(tc_b, W + 'bottom', {W + 'val': 'single', W + 'sz': '6'})
                ET.SubElement(ET.SubElement(ET.SubElement(tc, W + 'p'), W + 'r'), W + 't').text = f'列{r_idx}'
            sdt_c.append(tbl)

            # Chapter 2 via Word auto-numbering numPr + bracket non-citations [2024] & [0, 1] + theme font run + 图 10.1 / 图 1-1
            p_ch2 = ET.Element(W + 'p')
            ppr = ET.SubElement(p_ch2, W + 'pPr')
            ET.SubElement(ppr, W + 'pStyle', {W + 'val': 'Heading1'})
            npr = ET.SubElement(ppr, W + 'numPr')
            ET.SubElement(npr, W + 'ilvl', {W + 'val': '0'})
            ET.SubElement(npr, W + 'numId', {W + 'val': '10'})
            ET.SubElement(ET.SubElement(p_ch2, W + 'r'), W + 't').text = '实验结果'

            p_body = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_body, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            r_body = ET.SubElement(p_body, W + 'r')
            r_rpr = ET.SubElement(r_body, W + 'rPr')
            ET.SubElement(r_rpr, W + 'rFonts', {W + 'asciiTheme': 'minorHAnsi'})
            ET.SubElement(r_body, W + 't').text = (
                '如图 10.1 所示，数据集 [2024] 归一至区间 [0, 1]，孵化 3 小时；详见图 1-1。'
            )

            p_fig10 = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_fig10, W + 'pPr'), W + 'pStyle', {W + 'val': 'Caption'})
            r_cap10 = ET.SubElement(p_fig10, W + 'r')
            ET.SubElement(ET.SubElement(r_cap10, W + 'rPr'), W + 'b')
            ET.SubElement(r_cap10, W + 't').text = '图 10.1 '
            ET.SubElement(ET.SubElement(p_fig10, W + 'r'), W + 't').text = '多批整合（A）展示完整过程。'

            for offset, node in enumerate((sdt, p_ch2, p_body, p_fig10)):
                body.insert(insert_at + offset, node)

            # Acknowledgements section excluding first-person plural "我们"
            sect_idx = list(body).index(body.find(W + 'sectPr'))
            p_ack = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_ack, W + 'pPr'), W + 'pStyle', {W + 'val': 'Heading1'})
            ET.SubElement(ET.SubElement(p_ack, W + 'r'), W + 't').text = '致谢'
            p_ack_txt = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_ack_txt, W + 'r'), W + 't').text = '感谢实验室伙伴一路陪伴我们完成研究。'
            body.insert(sect_idx, p_ack)
            body.insert(sect_idx + 1, p_ack_txt)

        # Add numbering.xml (numId=10 -> 第2章) and theme1.xml (minorFont=Times New Roman)
        num_xml = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="2"/>'
            b'<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="\xe7\xac\xac%1\xe7\xab\xa0"/>'
            b'</w:lvl></w:abstractNum><w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
        )
        theme_xml = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ThesisTheme">'
            b'<a:themeElements><a:fontScheme name="Thesis"><a:majorFont><a:latin typeface="Times New Roman"/>'
            b'<a:ea typeface="SimHei"/><a:cs typeface="Times New Roman"/></a:majorFont>'
            b'<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface="SimSun"/>'
            b'<a:cs typeface="Times New Roman"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
        )
        modify(template, modern, {MAIN: build_modern_elements})
        with ZipFile(modern, 'a', ZIP_DEFLATED) as zf:
            if 'word/numbering.xml' not in zf.namelist():
                zf.writestr('word/numbering.xml', num_xml)
            if 'word/theme/theme1.xml' not in zf.namelist():
                zf.writestr('word/theme/theme1.xml', theme_xml)

        cfg = self.work / 'enable-checks.json'
        cfg.write_text(json.dumps({
            'audit': {'template_checks': ['Tbl-3Line', 'Acad-14', 'Err-7', 'Err-8']}
        }), encoding='utf-8')
        findings = self.audit(modern, '--template', template, '--citation-style', 'nature', '--profile-config', cfg)
        self.assertEqual(findings['Tbl-3Line']['status'], 'PASS')
        self.assertIn('识别 2 个内容表格', findings['Tbl-3Line']['summary'])
        self.assertEqual(findings['Fig-1']['status'], 'PASS')
        self.assertIn('图=2', findings['Fig-1']['summary'])
        self.assertEqual(findings['Fig-3-8']['status'], 'PASS')
        self.assertEqual(findings['Acad-14']['status'], 'PASS')
        self.assertEqual(findings['Err-7']['status'], 'PASS')
        self.assertEqual(findings['Err-8']['status'], 'PASS')
        self.assertEqual(findings['Ref-8']['status'], 'PASS')

    def test_populated_toc_missing_font_and_zero_ref_warnings(self):
        doc = self.work / 'toc-populated.docx'
        self.invoke('init', '--output', doc)
        defect = self.work / 'toc-defect.docx'

        def inject_toc_and_defects(root):
            body = root.find(W + 'body')
            blocks = list(body)
            ch1_i = next(i for i, el in enumerate(blocks) if '第一章' in visible(el))
            for idx_offset, label in enumerate(('第一章 绪论\t1', '第二章 材料与方法\t5', '致谢\t50')):
                p_toc = ET.Element(W + 'p')
                ET.SubElement(ET.SubElement(p_toc, W + 'pPr'), W + 'pStyle', {W + 'val': 'TOC1'})
                ET.SubElement(ET.SubElement(p_toc, W + 'r'), W + 't').text = label
                body.insert(ch1_i + idx_offset, p_toc)
            ref_i = next(i for i, el in enumerate(list(body)) if visible(el).strip() == '参考文献')
            p_bad = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_bad, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ET.SubElement(p_bad, W + 'r'), W + 't').text = '加入 10uL 试剂，基因 Tp53 上调。'
            body.insert(ref_i, p_bad)

        def strip_latin_fonts(styles_root):
            for rf in styles_root.iter(W + 'rFonts'):
                for attr in (W + 'ascii', W + 'hAnsi', W + 'asciiTheme', W + 'hAnsiTheme'):
                    rf.attrib.pop(attr, None)

        modify(doc, defect, {MAIN: inject_toc_and_defects, 'word/styles.xml': strip_latin_fonts})
        res = self.audit(defect, '--genes', 'Tp53', returncode=1)
        self.assertEqual(res['Err-4-5-6']['status'], 'FAIL')
        self.assertEqual(res['Err-8']['status'], 'FAIL')
        self.assertEqual(res['Acad-15-16']['status'], 'FAIL')

        no_refs = self.work / 'renamed-empty-refs.docx'
        def clear_refs(root):
            body = root.find(W + 'body')
            for el in list(body):
                if el.tag == W + 'p':
                    for r in el.findall(f'./{W}r'):
                        va = r.find(f'./{W}rPr/{W}vertAlign')
                        if va is not None and va.get(W + 'val') == 'superscript':
                            el.remove(r)
            blocks = list(body)
            ri = next(i for i, el in enumerate(blocks) if visible(el).strip() == '参考文献')
            blocks[ri].find(f'./{W}r/{W}t').text = '主要参考文献'
            end_i = next((i for i in range(ri + 1, len(blocks)) if visible(blocks[i]).startswith('个人简历')), len(blocks))
            for rem in blocks[ri + 1 : end_i]:
                body.remove(rem)

        modify(doc, no_refs, {MAIN: clear_refs})
        res_empty = self.audit(no_refs)
        for cid in ('Ref-2-7', 'Ref-8', 'Ref-9'):
            self.assertEqual(res_empty[cid]['status'], 'WARN')

    def test_repair_tables_and_leading_spaces_safety(self):
        draft = self.work / 'rich-draft.docx'
        template = self.work / 'rich-template.docx'
        self.invoke('init', '--output', template)
        self.invoke('init', '--output', draft, '--title-zh', '水稻抗病基因研究')

        def customize_draft(root):
            body = root.find(W + 'body')
            ref_el = next(el for el in list(body) if visible(el).strip() == '参考文献')
            ref_idx = list(body).index(ref_el)

            # Paragraph starting with hyperlink followed by space-leading run
            p_link = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_link, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            hl = ET.SubElement(p_link, W + 'hyperlink')
            ET.SubElement(ET.SubElement(hl, W + 'r'), W + 't').text = 'GeneA'
            ET.SubElement(ET.SubElement(p_link, W + 'r'), W + 't').text = ' plays a key role.'

            # Paragraph explicitly overriding firstLine=0
            p_zero = ET.Element(W + 'p')
            ppr_z = ET.SubElement(p_zero, W + 'pPr')
            ET.SubElement(ppr_z, W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ppr_z, W + 'ind', {W + 'firstLine': '0'})
            ET.SubElement(ET.SubElement(p_zero, W + 'r'), W + 't').text = '  intentional align'

            # Indented paragraph with stray leading spaces
            p_stray = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_stray, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ET.SubElement(p_stray, W + 'r'), W + 't').text = '  \u3000正文多余空格'

            # Two-row header table inside <w:sdt>
            sdt = ET.Element(W + 'sdt')
            tbl = ET.SubElement(ET.SubElement(sdt, W + 'sdtContent'), W + 'tbl')
            for ri in range(3):
                tr = ET.SubElement(tbl, W + 'tr')
                if ri < 2:
                    ET.SubElement(ET.SubElement(tr, W + 'trPr'), W + 'tblHeader')
                ET.SubElement(ET.SubElement(ET.SubElement(tr, W + 'tc'), W + 'p'), W + 'r')

            for off, el in enumerate((p_link, p_zero, p_stray, sdt)):
                body.insert(ref_idx + off, el)

        sample = self.work / 'rich-sample.docx'
        modify(draft, sample, {MAIN: customize_draft})
        with ZipFile(sample) as zf:
            cover_before = ET.tostring(next(ET.fromstring(zf.read(MAIN)).iter(W + 'tbl')))

        repaired = self.work / 'rich-repaired.docx'
        self.invoke('repair', '--docx', sample, '--template', template, '--output', repaired,
                    '--fix-tables', '--clean-leading-spaces')
        with ZipFile(repaired) as zf:
            root_after = ET.fromstring(zf.read(MAIN))
        cover_after = ET.tostring(next(root_after.iter(W + 'tbl')))
        self.assertEqual(cover_before, cover_after)
        txt_after = '\n'.join(visible(p) for p in root_after.iter(W + 'p'))
        self.assertIn('GeneA plays a key role.', txt_after)
        self.assertIn('  intentional align', txt_after)
        self.assertIn('正文多余空格', txt_after)
        self.assertNotIn('  \u3000正文多余空格', txt_after)
        sdt_tbl = root_after.find(f'.//{W}sdt//{W}tbl')
        self.assertEqual(sdt_tbl.find(f'./{W}tblPr/{W}tblBorders/{W}top').get(W + 'sz'), '12')

    def test_repair_missing_settings_and_styleref_plus_page_footer(self):
        template = self.work / 'base-tpl.docx'
        self.invoke('init', '--output', template)
        draft = self.work / 'broken-4sec.docx'

        def make_4sec(root, footer_rel_id):
            body = root.find(W + 'body')
            sects = list(root.iter(W + 'sectPr'))
            # Make S2 (index 1) decimal without start, and insert a Body2 section S4
            s2_num = sects[1].find(W + 'pgNumType')
            if s2_num is not None:
                s2_num.set(W + 'fmt', 'decimal')
                s2_num.attrib.pop(W + 'start', None)
            ref_p = next(i for i, el in enumerate(list(body)) if visible(el).strip() == '参考文献')
            p_break = ET.Element(W + 'p')
            ppr = ET.SubElement(p_break, W + 'pPr')
            s_new = ET.SubElement(ppr, W + 'sectPr')
            ET.SubElement(s_new, W + 'footerReference', {W + 'type': 'default', R + 'id': footer_rel_id})
            ET.SubElement(s_new, W + 'pgSz', {W + 'w': '11906', W + 'h': '16838'})
            ET.SubElement(s_new, W + 'pgMar', {
                W + 'top': '1701', W + 'right': '1701', W + 'bottom': '1417',
                W + 'left': '1701', W + 'header': '851', W + 'footer': '992', W + 'gutter': '0',
            })
            ET.SubElement(s_new, W + 'pgNumType', {W + 'fmt': 'decimal', W + 'start': '1'})
            body.insert(ref_p, p_break)

        with ZipFile(template) as zsrc, ZipFile(draft, 'w', ZIP_DEFLATED) as zdst:
            rels_root = ET.fromstring(zsrc.read('word/_rels/document.xml.rels'))
            s2_footer_id = None
            doc_tmp = ET.fromstring(zsrc.read(MAIN))
            sects_tmp = list(doc_tmp.iter(W + 'sectPr'))
            f_ref = sects_tmp[1].find(W + 'footerReference')
            s2_footer_id = f_ref.get(R + 'id')
            s2_footer_target = 'word/' + next(r.get('Target') for r in rels_root if r.get('Id') == s2_footer_id).lstrip('/')
            for name in zsrc.namelist():
                if name == 'word/settings.xml':
                    continue
                data = zsrc.read(name)
                if name == MAIN:
                    r_doc = ET.fromstring(data)
                    make_4sec(r_doc, s2_footer_id)
                    data = ET.tostring(r_doc, encoding='utf-8', xml_declaration=True)
                elif name == s2_footer_target:
                    r_ft = ET.fromstring(data)
                    p = r_ft.find(W + 'p')
                    if p is None:
                        p = ET.SubElement(r_ft, W + 'p')
                    # Precede PAGE with STYLEREF 1 complex field
                    for item in list(p):
                        p.remove(item)
                    ET.SubElement(ET.SubElement(p, W + 'r'), W + 'fldChar', {W + 'fldCharType': 'begin'})
                    ET.SubElement(ET.SubElement(p, W + 'r'), W + 'instrText').text = ' STYLEREF 1 '
                    ET.SubElement(ET.SubElement(p, W + 'r'), W + 'fldChar', {W + 'fldCharType': 'end'})
                    ET.SubElement(ET.SubElement(p, W + 'r'), W + 'fldChar', {W + 'fldCharType': 'begin'})
                    ET.SubElement(ET.SubElement(p, W + 'r'), W + 'instrText').text = ' PAGE '
                    ET.SubElement(ET.SubElement(p, W + 'r'), W + 'fldChar', {W + 'fldCharType': 'end'})
                    data = ET.tostring(r_ft, encoding='utf-8', xml_declaration=True)
                zdst.writestr(name, data)

        repaired = self.work / 'fixed-4sec.docx'
        self.invoke('repair', '--docx', draft, '--template', template, '--output', repaired)
        after = self.audit(repaired, '--template', template)
        self.assertEqual(after['Layout-Settings']['status'], 'PASS')
        self.assertEqual(after['Layout-1']['status'], 'PASS')
        self.assertEqual(after['Layout-SecNum']['status'], 'PASS')

    def test_rigorous_default_rules_and_author_cv(self):
        final_doc = self.work / 'rigorous-final.docx'
        self.invoke('init', '--output', final_doc, '--mode', 'phd-thesis', '--review-stage', 'final',
                    '--university', '示例大学', '--school', '生科院', '--author', '张科研', '--major', '生物学')
        ok_findings = self.audit(final_doc)
        for cid in ('Acad-14', 'Acad-CV', 'Ref-2-7', 'Err-7', 'Err-2-11-12', 'Err-14'):
            self.assertEqual(ok_findings[cid]['status'], 'PASS', cid)
        self.assertEqual(ok_findings['Fig-5']['status'], 'INFO')
        with ZipFile(final_doc) as zf:
            doc_txt = visible(ET.fromstring(zf.read(MAIN)))
        self.assertIn('张科研', doc_txt)
        self.assertIn('示例大学', doc_txt)

        # DrawingML & SVG text < 6号字 (7.5 pt) must fail Fig-5; >= 7.5 pt must pass Fig-5
        A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        ASVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
        def add_vector_fig(root, draw_pt_100: str):
            body = root.find(W + 'body')
            ch1_p = next(el for el in body if el.tag == W + 'p' and '第一章' in visible(el))
            p = ET.Element(W + 'p')
            r = ET.SubElement(p, W + 'r')
            d = ET.SubElement(r, W + 'drawing')
            ar = ET.SubElement(d, f'{{{A_NS}}}r')
            ET.SubElement(ar, f'{{{A_NS}}}rPr', {'sz': draw_pt_100})
            ET.SubElement(ar, f'{{{A_NS}}}t').text = 'ScaleLabel'
            ET.SubElement(d, f'{{{ASVG_NS}}}svgBlip', {R + 'embed': 'rIdSvg1'})
            body.insert(list(body).index(ch1_p) + 2, p)

        def add_svg_rel(rels_root):
            ET.SubElement(rels_root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', {
                'Id': 'rIdSvg1',
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                'Target': 'media/fig_test.svg',
            })

        small_fig = self.work / 'small-fig-font.docx'
        modify(final_doc, small_fig, {MAIN: lambda r: add_vector_fig(r, '600'), 'word/_rels/document.xml.rels': add_svg_rel})
        with ZipFile(small_fig, 'a', ZIP_DEFLATED) as zf:
            zf.writestr('word/media/fig_test.svg', '<svg xmlns="http://www.w3.org/2000/svg"><text font-size="6.5pt">TinyAxis</text></svg>')
        res_small = self.audit(small_fig, returncode=1)
        self.assertEqual(res_small['Fig-5']['status'], 'FAIL')
        self.assertEqual(len(res_small['Fig-5']['details']), 2)

        ok_fig = self.work / 'ok-fig-font.docx'
        modify(final_doc, ok_fig, {MAIN: lambda r: add_vector_fig(r, '800'), 'word/_rels/document.xml.rels': add_svg_rel})
        with ZipFile(ok_fig, 'a', ZIP_DEFLATED) as zf:
            zf.writestr('word/media/fig_test.svg', '<svg xmlns="http://www.w3.org/2000/svg"><text font-size="8pt">ClearAxis</text></svg>')
        self.assertEqual(self.audit(ok_fig)['Fig-5']['status'], 'PASS')

        # Ahead-of-print DOI passes, published DOI fails; mixed Ch2 time units fail; singular "我" fails ("我国" passes); vendor/subpanel commas fail; CJK half-space fails; missing CV fails
        def inject_rigorous_defects(root):
            body = root.find(W + 'body')
            ch2_p = next(el for el in body if el.tag == W + 'p' and visible(el).strip().startswith('第二章'))
            ch2_idx = list(body).index(ch2_p)
            p1 = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p1, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ET.SubElement(p1, W + 'r'), W + 't').text = (
                '我通过实验验证，我国学者（Smith et al., 2020）采用设备（P-2000, Sutter Instrument）与仪器（PCR仪、Bio-Rad）及图 1.1 A, B 完成测试。'
            )
            body.insert(ch2_idx + 1, p1)

            p2 = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p2, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ET.SubElement(p2, W + 'r'), W + 't').text = (
                '细胞 培养后以 50\u3000mL 缓冲液洗涤，在 37°C 孵育 15 分钟并离心 30 s（参照图 1.1）'
            )
            r_sup = ET.SubElement(p2, W + 'r')
            ET.SubElement(ET.SubElement(r_sup, W + 'rPr'), W + 'vertAlign', {W + 'val': 'superscript'})
            ET.SubElement(r_sup, W + 't').text = '3,4'
            ET.SubElement(ET.SubElement(p2, W + 'r'), W + 't').text = '。'
            body.insert(ch2_idx + 2, p2)

            # Inject ref 3 (published with DOI -> FAIL) and ref 4 (ahead-of-print with DOI -> PASS)
            cv_h = next(el for el in list(body) if visible(el).strip().startswith('个人简历'))
            cv_pos = list(body).index(cv_h)
            for off, (prefix, jrn, vol, tail) in enumerate((
                ('3. Alpha, A. Published work. ', 'Nat. Commun.', '15', ', 4321 (2024). https://doi.org/10.1038/s41467-024-00001-x'),
                ('4. Beta, B. Preprint work. ', 'Nature', '', ' https://doi.org/10.1038/s41586-026-00001-x (2026)'),
            )):
                rp = ET.Element(W + 'p')
                ET.SubElement(ET.SubElement(rp, W + 'pPr'), W + 'pStyle', {W + 'val': 'BibEntry'})
                ET.SubElement(ET.SubElement(rp, W + 'r'), W + 't').text = prefix
                rj = ET.SubElement(rp, W + 'r')
                ET.SubElement(ET.SubElement(rj, W + 'rPr'), W + 'i')
                ET.SubElement(rj, W + 't').text = jrn
                if vol:
                    ET.SubElement(ET.SubElement(rp, W + 'r'), W + 't').text = ' '
                    rb = ET.SubElement(rp, W + 'r')
                    ET.SubElement(ET.SubElement(rb, W + 'rPr'), W + 'b')
                    ET.SubElement(rb, W + 't').text = vol
                ET.SubElement(ET.SubElement(rp, W + 'r'), W + 't').text = tail
                body.insert(cv_pos + off, rp)

            # Remove CV block completely to verify Acad-CV fails when omitted
            ack_h = next(el for el in list(body) if visible(el).strip().replace('\u3000', '').replace(' ', '') == '致谢')
            cur_list = list(body)
            for rem in cur_list[cur_list.index(cv_h) : cur_list.index(ack_h)]:
                body.remove(rem)

        defect_doc = self.work / 'rigorous-defects.docx'
        modify(final_doc, defect_doc, {MAIN: inject_rigorous_defects})
        defects = self.audit(defect_doc, returncode=1)
        for cid in ('Acad-14', 'Acad-CV', 'Ref-2-7', 'Err-7', 'Err-2-11-12', 'Err-14'):
            self.assertEqual(defects[cid]['status'], 'FAIL', f'{cid}: {defects[cid]}')
        self.assertEqual(len(defects['Ref-2-7']['details']), 1)
        self.assertIn('文献 3', defects['Ref-2-7']['details'][0])

        # Custom --profile-config can explicitly override/disable rules when a user supplies school differences
        custom_override = self.work / 'disable-rigorous.json'
        custom_override.write_text(json.dumps({
            'audit': {
                'forbid_published_doi': False,
                'disabled_checks': ['Acad-14', 'Acad-CV', 'Err-7', 'Err-2-11-12', 'Err-14'],
            }
        }), encoding='utf-8')
        relaxed = self.audit(defect_doc, '--profile-config', custom_override)
        for cid in ('Acad-14', 'Acad-CV', 'Err-7', 'Err-2-11-12', 'Err-14'):
            self.assertEqual(relaxed[cid]['status'], 'INFO', cid)
        self.assertEqual(relaxed['Ref-2-7']['status'], 'PASS')

        # Valid biomedical terms & enumerations must NOT false-trigger Acad-14, Err-2-11-12, Err-7, or Err-14
        def inject_false_positive_guards(root):
            body = root.find(W + 'body')
            ch2_p = next(el for el in body if el.tag == W + 'p' and visible(el).strip().startswith('第二章'))
            ch2_idx = list(body).index(ch2_p)
            p_ok = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_ok, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ET.SubElement(p_ok, W + 'r'), W + 't').text = (
                '本论文纳入样本人数为 120 例，比较东亚及日本人群队列。'
                '实验使用DNA、RNA提取试剂（含裂解液、洗脱液等试剂）、生物材料（细胞、组织等生物样本）与设备（Thermo、Sigma 等供应商）。'
                '参考 1990s 的方案并观察 30s 核糖体亚基后，在 37°C 孵育 15 分钟。'
            )
            body.insert(ch2_idx + 1, p_ok)

        guard_doc = self.work / 'rigorous-fp-guards.docx'
        modify(final_doc, guard_doc, {MAIN: inject_false_positive_guards})
        guard_res = self.audit(guard_doc)
        for cid in ('Acad-14', 'Err-7', 'Err-2-11-12', 'Err-14'):
            self.assertEqual(guard_res[cid]['status'], 'PASS', f'{cid}: {guard_res[cid]}')

        # Inline deleted run & textbox with half-width CJK spacing must be ignored by Err-14
        def inject_del_and_tbx(root):
            body = root.find(W + 'body')
            ch2_p = next(el for el in body if el.tag == W + 'p' and visible(el).strip().startswith('第二章'))
            p_del = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(p_del, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            w_del = ET.SubElement(p_del, W + 'del', {W + 'id': '99', W + 'author': 'reviewer'})
            ET.SubElement(ET.SubElement(w_del, W + 'r'), W + 'delText').text = '已删 文本'
            ET.SubElement(ET.SubElement(w_del, W + 'r'), W + 't').text = '已删 文本'
            body.insert(list(body).index(ch2_p) + 1, p_del)

        guard_del_doc = self.work / 'rigorous-del-guard.docx'
        modify(final_doc, guard_del_doc, {MAIN: inject_del_and_tbx})
        del_res = self.audit(guard_del_doc, returncode=1)
        self.assertEqual(del_res['Pkg-Clean']['status'], 'FAIL')
        self.assertEqual(del_res['Err-14']['status'], 'PASS')

        # Corrupt embedded SVG and inherited <g font-size="6pt"> must fail Fig-5 explicitly
        bad_svg_doc = self.work / 'bad-svg.docx'
        modify(final_doc, bad_svg_doc, {MAIN: lambda r: add_vector_fig(r, '800'), 'word/_rels/document.xml.rels': add_svg_rel})
        with ZipFile(bad_svg_doc, 'a', ZIP_DEFLATED) as zf:
            zf.writestr('word/media/fig_test.svg', '<svg><unclosed>')
        self.assertEqual(self.audit(bad_svg_doc, returncode=1)['Fig-5']['status'], 'FAIL')

        nested_svg_doc = self.work / 'nested-small-svg.docx'
        modify(final_doc, nested_svg_doc, {MAIN: lambda r: add_vector_fig(r, '800'), 'word/_rels/document.xml.rels': add_svg_rel})
        with ZipFile(nested_svg_doc, 'a', ZIP_DEFLATED) as zf:
            zf.writestr('word/media/fig_test.svg', '<svg xmlns="http://www.w3.org/2000/svg"><g font-size="6pt"><text>InheritedSmall</text></g></svg>')
        self.assertEqual(self.audit(nested_svg_doc, returncode=1)['Fig-5']['status'], 'FAIL')

        # Published paper containing "preprint" in title with colon article numbering must still fail Ref-2-7 if DOI included
        def inject_preprint_title_published_doi(root):
            body = root.find(W + 'body')
            cv_h = next(el for el in list(body) if visible(el).strip().startswith('个人简历'))
            rp = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(rp, W + 'pPr'), W + 'pStyle', {W + 'val': 'BibEntry'})
            ET.SubElement(ET.SubElement(rp, W + 'r'), W + 't').text = '3. Gamma, G. Preprint benchmark evaluation. '
            rj = ET.SubElement(rp, W + 'r')
            ET.SubElement(ET.SubElement(rj, W + 'rPr'), W + 'i')
            ET.SubElement(rj, W + 't').text = 'Nat. Commun.'
            ET.SubElement(ET.SubElement(rp, W + 'r'), W + 't').text = ' '
            rb = ET.SubElement(rp, W + 'r')
            ET.SubElement(ET.SubElement(rb, W + 'rPr'), W + 'b')
            ET.SubElement(rb, W + 't').text = '15'
            ET.SubElement(ET.SubElement(rp, W + 'r'), W + 't').text = ': 4321 (2024). https://doi.org/10.1038/s41467-024-00002-y'
            body.insert(list(body).index(cv_h), rp)
            ch1_p = next(el for el in body if el.tag == W + 'p' and '第一章' in visible(el))
            cite_p = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(cite_p, W + 'pPr'), W + 'pStyle', {W + 'val': 'BodyTextZH'})
            ET.SubElement(ET.SubElement(cite_p, W + 'r'), W + 't').text = '先前工作已描述'
            rs = ET.SubElement(cite_p, W + 'r')
            ET.SubElement(ET.SubElement(rs, W + 'rPr'), W + 'vertAlign', {W + 'val': 'superscript'})
            ET.SubElement(rs, W + 't').text = '3'
            ET.SubElement(ET.SubElement(cite_p, W + 'r'), W + 't').text = '。'
            body.insert(list(body).index(ch1_p) + 1, cite_p)

        preprint_pub = self.work / 'preprint-title-pub-doi.docx'
        modify(final_doc, preprint_pub, {MAIN: inject_preprint_title_published_doi})
        self.assertEqual(self.audit(preprint_pub, returncode=1)['Ref-2-7']['status'], 'FAIL')

        # Blind stage without --sensitive-names must emit WARN (not silent PASS), and still FAIL on CV leaks
        blind_doc = self.work / 'rigorous-blind.docx'
        self.invoke('init', '--output', blind_doc, '--mode', 'phd-thesis', '--review-stage', 'blind',
                    '--university', '示例大学', '--school', '生科院', '--author', '张科研', '--advisor', '王导师')
        self.assertEqual(self.audit(blind_doc, '--review-stage', 'blind')['Acad-10-Blind']['status'], 'WARN')
        self.assertEqual(self.audit(blind_doc, '--review-stage', 'blind', '--sensitive-names', '张科研,王导师')['Acad-10-Blind']['status'], 'PASS')
        leaked_cv = self.work / 'leaked-cv-blind.docx'
        def unredact_cv(root):
            p = next(el for el in root.iter(W + 'p') if '匿名评审版本隐去姓名' in visible(el))
            p.find(f'.//{W}t').text = '姓名：外漏研究者'
        modify(blind_doc, leaked_cv, {MAIN: unredact_cv})
        leaked_res = self.audit(leaked_cv, '--review-stage', 'blind', returncode=1)
        self.assertEqual(leaked_res['Acad-CV']['status'], 'FAIL')
        self.assertEqual(leaked_res['Acad-10-Blind']['status'], 'FAIL')


if __name__ == '__main__':
    unittest.main()
