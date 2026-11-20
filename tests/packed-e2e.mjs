import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const python = process.env.PYTHON || 'python';
const packageName = 'dsh-plugin-bio-thesis';

function run(program, args, cwd, expected = 0) {
  const result = spawnSync(program, args, { cwd, encoding: 'utf8', timeout: 120_000 });
  if (result.error) throw result.error;
  assert.equal(result.status, expected, `${program} ${args.join(' ')}\n${result.stdout}\n${result.stderr}`);
  return result.stdout;
}

function npm(args, cwd) {
  // Via npm scripts or a Windows Node distribution, run npm without a command shell.
  const cli = process.env.npm_execpath || (process.platform === 'win32'
    ? path.join(path.dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js')
    : undefined);
  return cli ? run(process.execPath, [cli, ...args], cwd) : run('npm', args, cwd);
}

function zip(args, cwd) {
  return run(python, ['-B', '-X', 'utf8', '-m', 'zipfile', ...args], cwd);
}

function checkSummary(report) {
  const expected = { PASS: 0, WARN: 0, FAIL: 0, INFO: 0 };
  for (const check of report.checks) expected[check.status] += 1;
  assert.deepEqual(report.review_summary, expected);
  return expected;
}

// A small independent school template. The title spans two Word text runs.
// The landscape Letter section and custom title style make preset substitution observable.
async function makeSchoolTemplate(folder, destination) {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Override PartName="/word/document.xml" ContentType="${contentType}.template.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="${contentType}.styles+xml"/>
      <Override PartName="/word/header1.xml" ContentType="${contentType}.header+xml"/>
    </Types>`,
    '_rels/.rels': `<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="${rel}">
      <Relationship Id="style" Type="${r}/styles" Target="styles.xml"/>
      <Relationship Id="hdr" Type="${r}/header" Target="header1.xml"/>
    </Relationships>`,
    'word/header1.xml': `<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>Example Graduate School</w:t></w:r></w:p></w:hdr>`,
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:styleId="SchoolTitle">
      <w:name w:val="School Title"/><w:rPr><w:b/><w:sz w:val="46"/></w:rPr></w:style></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="SchoolTitle"/></w:pPr><w:r><w:t>{{title_</w:t></w:r><w:r><w:t>zh}}</w:t></w:r></w:p>
      <w:p><w:r><w:t>{{author}}</w:t></w:r></w:p>
      <w:sectPr><w:headerReference w:type="default" r:id="hdr"/>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="1800" w:right="1000" w:bottom="1300" w:left="1100" w:header="550" w:footer="500" w:gutter="0"/>
      </w:sectPr></w:body></w:document>`,
  };
  for (const [name, value] of Object.entries(parts)) {
    const output = path.join(folder, name);
    await fs.mkdir(path.dirname(output), { recursive: true });
    await fs.writeFile(output, value);
  }
  zip(['-c', destination, '[Content_Types].xml', '_rels', 'word'], folder);
}

test('installed tarball can generate and audit a thesis using only its shipped resources', async t => {
  const tempRoot = path.resolve(os.tmpdir());
  const work = await fs.mkdtemp(path.join(tempRoot, 'thesis-package-e2e-'));
  t.after(async () => {
    assert.equal(path.dirname(work), tempRoot);
    assert.match(path.basename(work), /^thesis-package-e2e-[a-z0-9]+$/i);
    await fs.rm(work, { recursive: true, force: true });
  });
  const consumer = path.join(work, 'consumer');
  await fs.mkdir(consumer);
  await fs.writeFile(path.join(consumer, 'package.json'), '{"name":"test-consumer","private":true}');
  const artifact = JSON.parse(npm(['pack', '--json', '--ignore-scripts', '--offline',
    '--pack-destination', work, '--cache', path.join(work, 'npm-cache')], root))[0];
  assert.equal(artifact.name, packageName);
  npm(['install', path.join(work, artifact.filename), '--offline', '--ignore-scripts', '--legacy-peer-deps',
    '--no-audit', '--no-fund', '--cache', path.join(work, 'npm-cache')], consumer);

  const resolve = createRequire(path.join(consumer, 'package.json')).resolve;
  const installed = await fs.realpath(path.dirname(resolve(`${packageName}/package.json`)));
  assert.equal(installed, await fs.realpath(path.join(consumer, 'node_modules', packageName)));
  const plugin = await import(pathToFileURL(resolve(packageName)).href);
  const providers = [];
  plugin.apply({ skills: { registerProvider: factory => providers.push(factory()) } });
  assert.equal(providers.length, 1);
  const [candidate] = await providers[0].list();
  const skill = await providers[0].get(candidate);
  assert.equal(skill.name, 'bio-thesis-formatter');
  assert.match(skill.content, /bio_thesis\.py/);
  const skillDir = path.join(installed, 'skills', skill.name);
  assert.equal(await fs.realpath(skill.resourceBase.path), await fs.realpath(skillDir));
  const cli = (args, expected) => run(python, ['-B', '-X', 'utf8',
    path.join(skillDir, 'scripts', 'bio_thesis.py'), ...args], consumer, expected);

  await t.test('generic draft, shipped presets and checklist are usable after install', async () => {
    const presets = JSON.parse(cli(['list-presets']));
    assert.deepEqual(presets.map(p => p.preset).sort(), ['cn-gb7714-general', 'cn-lifesci-default']);
    cli(['init', '--output', 'generic.docx']);
    assert.match(cli(['audit', '--docx', 'generic.docx']), /AUDIT_SUMMARY=FAIL:0\b/);
    cli(['export-checklist', '--output', 'checklist.docx']);
    assert.deepEqual(await fs.readFile(path.join(consumer, 'checklist.docx')),
      await fs.readFile(path.join(skillDir, 'resources/bio-thesis-self-check-65.docx')));
  });

  const source = path.join(consumer, 'school.dotx');
  const originalParts = path.join(consumer, 'template-parts');
  const generatedParts = path.join(consumer, 'generated-parts');
  const output = path.join(consumer, 'draft.docx');
  const identity = 'Synthetic Reviewer Example';
  const blindAudit = ['--template', source, '--review-stage', 'blind', '--sensitive-names', identity];
  await makeSchoolTemplate(originalParts, source);
  const originalBytes = await fs.readFile(source);

  await t.test('external DOTX controls the draft, split title and masked author', async () => {
    cli(['init', '--template', source, '--output', output, '--title-zh', '跨学科样例',
      '--review-stage', 'blind', '--author', identity]);
    assert.deepEqual(await fs.readFile(source), originalBytes);
    zip(['-e', output, generatedParts], consumer);
    const document = await fs.readFile(path.join(generatedParts, 'word/document.xml'), 'utf8');
    assert.match(document, /跨学科样例/);
    assert.match(document, /匿名评审隐去/);
    assert.ok(!document.includes(identity) && !document.includes('{{'));
    assert.match(document, /w:orient="landscape"/);
    assert.match(document, /w:top="1800"/);
    assert.match(await fs.readFile(path.join(generatedParts, '[Content_Types].xml'), 'utf8'),
      /wordprocessingml\.document\.main\+xml/);
    for (const part of ['word/styles.xml', 'word/header1.xml', 'word/_rels/document.xml.rels', '_rels/.rels']) {
      assert.deepEqual(await fs.readFile(path.join(generatedParts, part)), await fs.readFile(path.join(originalParts, part)));
    }
    cli(['audit', '--docx', output, ...blindAudit, '--fail-on-warn', '--json-report', 'good.json', '--md-report', 'good.md']);
    const report = JSON.parse(await fs.readFile(path.join(consumer, 'good.json'), 'utf8'));
    const summary = checkSummary(report);
    assert.ok(summary.INFO > 0 && summary.WARN === 0 && summary.FAIL === 0);
    const tasks = report.manual_review;
    assert.deepEqual(tasks.map(task => task.id).sort(),
      ['blind-images-signatures', 'refresh-fields', 'rendered-figures-tables', 'rendered-pages']);
    assert.ok(tasks.every(task => task.status === 'PENDING' && !task.action.includes(identity)));
    const md = await fs.readFile(path.join(consumer, 'good.md'), 'utf8');
    for (const task of tasks) assert.ok(md.includes('- [ ] ' + task.action));
    for (const [status, count] of Object.entries(summary)) assert.ok(md.includes(status + ' ' + count));
    assert.ok(md.includes(report.report_notice));
    for (const id of ['Layout-1', 'Layout-TemplateStyles', 'Layout-SecNum', 'Acad-10-Blind']) {
      assert.equal(report.checks.find(c => c.check_id === id)?.status, 'PASS', id);
    }
  });

  await t.test('a template style warning can stop the build without changing the reported findings', async () => {
    const warnParts = path.join(consumer, 'warn-parts');
    zip(['-e', output, warnParts], consumer);
    const styles = path.join(warnParts, 'word/styles.xml');
    const xml = await fs.readFile(styles, 'utf8');
    assert.match(xml, /w:sz w:val="46"/);
    await fs.writeFile(styles, xml.replace('w:sz w:val="46"', 'w:sz w:val="30"'));
    const doc = path.join(consumer, 'warn.docx');
    zip(['-c', doc, '[Content_Types].xml', '_rels', 'word'], warnParts);
    const audit = ['audit', '--docx', doc, '--template', source];
    cli([...audit, '--json-report', 'warn.json']);
    cli([...audit, '--fail-on-warn', '--json-report', 'strict.json'], 1);
    const standard = JSON.parse(await fs.readFile(path.join(consumer, 'warn.json'), 'utf8'));
    const strict = JSON.parse(await fs.readFile(path.join(consumer, 'strict.json'), 'utf8'));
    assert.deepEqual(standard.checks, strict.checks);
    assert.deepEqual(standard.review_summary, strict.review_summary);
    const summary = checkSummary(strict);
    assert.equal(summary.FAIL, 0);
    assert.equal(summary.WARN, 1);
    assert.deepEqual(strict.checks.filter(c => c.status === 'WARN').map(c => c.check_id), ['Layout-TemplateStyles']);
    assert.deepEqual(strict.manual_review.map(task => task.id).sort(),
      ['refresh-fields', 'rendered-figures-tables', 'rendered-pages']);
  });

  await t.test('drift and a supplied identity reintroduced into the draft both fail', async () => {
    const document = path.join(generatedParts, 'word/document.xml');
    const xml = await fs.readFile(document, 'utf8');
    await fs.writeFile(document, xml.replace('w:top="1800"', 'w:top="720"').replace('跨学科样例', identity));
    const damaged = path.join(consumer, 'damaged.docx');
    zip(['-c', damaged, '[Content_Types].xml', '_rels', 'word'], generatedParts);
    cli(['audit', '--docx', damaged, ...blindAudit, '--json-report', 'bad.json'], 1);
    const report = JSON.parse(await fs.readFile(path.join(consumer, 'bad.json'), 'utf8'));
    assert.equal(checkSummary(report).FAIL, 2);
    assert.deepEqual(report.checks.filter(c => c.status === 'FAIL').map(c => c.check_id).sort(),
      ['Acad-10-Blind', 'Layout-1']);
  });

  await t.test('input and report collisions preserve original bytes', async () => {
    const before = await fs.readFile(output);
    cli(['init', '--template', output, '--output', output, '--overwrite'], 1);
    cli(['audit', '--docx', output, '--fail-on-warn', '--json-report', output], 2);
    assert.deepEqual(await fs.readFile(output), before);
    assert.deepEqual(await fs.readFile(source), originalBytes);
  });
});
