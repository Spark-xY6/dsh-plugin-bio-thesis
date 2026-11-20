// Run from a checkout, with DSH_RUNTIME pointing to a local npm install prefix.
import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync } from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { dirname, isAbsolute, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const plugin = 'dsh-plugin-bio-thesis'
const skill = 'bio-thesis-formatter'
const profile = 'thesis-evaluation'
const runtime = process.env.DSH_RUNTIME
const spec = process.env.DSH_PLUGIN_SPEC
assert.ok(runtime, 'Set DSH_RUNTIME to the npm prefix containing @deepseek-ai/dsh')
assert.ok(spec && ((isAbsolute(spec) && spec.endsWith('.tgz')) || /^github:[\w.-]+\/[\w.-]+#[0-9a-f]{40}$/i.test(spec)),
  'Set DSH_PLUGIN_SPEC to an absolute .tgz or github:owner/repo#<full-commit-sha>')
const fromPrefix = createRequire(join(resolve(runtime), 'package.json'))
const hostManifest = fromPrefix.resolve('@deepseek-ai/dsh/package.json')
const fromHost = createRequire(hostManifest)
const host = JSON.parse(readFileSync(hostManifest, 'utf8'))
const bin = join(dirname(hostManifest), host.bin.dsh)
const python = process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3')
const testsDir = dirname(fileURLToPath(import.meta.url))
const initialDirectory = process.cwd()
const parent = realpathSync(process.env.DSH_E2E_PARENT || tmpdir())
const work = mkdtempSync(join(parent, 'dsh-thesis-'))
const home = join(work, 'home')
const workspace = join(work, 'documents')
mkdirSync(workspace)
process.env.DSH_HOME = home
process.env.DSH_TELEMETRY_DISABLED = '1'
process.env.CI = 'true'
process.chdir(workspace)

function command(executable, args, expected = 0) {
  const result = spawnSync(executable, args, { cwd: workspace, encoding: 'utf8', windowsHide: true,
    timeout: 180_000, maxBuffer: 16 * 1024 * 1024 })
  if (result.error) throw result.error
  assert.equal(result.status, expected, `${executable} exited with ${result.status}\n${result.stdout}\n${result.stderr}`)
  return result.stdout
}

function outcome(file, id) {
  const report = JSON.parse(readFileSync(join(workspace, file), 'utf8'))
  const check = report.checks.find(item => item.check_id === id)
  assert.ok(check, `Missing check ${id}`)
  return check.status
}

let shutdown
let passed = false
try {
  command(process.execPath, [bin, 'plugin', '--profile', profile, 'add', spec])
  const profileDir = join(home, 'profiles', profile)
  const installed = JSON.parse(readFileSync(join(profileDir, 'package.json'), 'utf8'))
  assert.ok(installed.dsh.profile.bundles.includes(plugin))
  const dump = command(process.execPath, [bin, '--profile', profile, '--dump-config'])
  assert.ok(dump.includes(plugin) && dump.includes('skill-bio-thesis'))

  // Use exactly the profile boot and environment loader used by the installed CLI.
  const { runProfile } = await import(pathToFileURL(fromHost.resolve('@deepseek-ai/dsh/profile-boot')))
  const { loadLayeredEnv } = await import(pathToFileURL(fromHost.resolve('@deepseek-ai/dsh-app-boot')))
  const boot = await runProfile({ environment: loadLayeredEnv('dsh'), profile, patchFiles: [], args: [] })
  shutdown = boot.shutdown
  const options = { cwd: workspace }
  const listed = (await boot.ctx.skills.list(options)).find(item => item.name === skill)
  assert.equal(listed?.provider, skill)
  assert.equal(listed.source, 'bundled')
  assert.ok(listed.invocation.modelInvocable && listed.invocation.userInvocable)
  const definition = await boot.ctx.skills.get(skill, options)
  assert.equal(definition.resourceBase.kind, 'directory')
  const base = realpathSync(definition.resourceBase.path)
  const packageRoot = realpathSync(join(profileDir, 'node_modules', plugin))
  assert.equal(base, realpathSync(join(packageRoot, 'skills', skill)))
  assert.equal(realpathSync(definition.path), realpathSync(join(base, 'SKILL.md')))
  assert.ok(definition.content.includes('bio_thesis.py'))
  console.log(`DSH_HOST=PASS version=${host.version} provider=${listed.provider} source=${listed.source}`)

  // Reuse the synthetic fixture builder, with the master resource supplied by the installed skill.
  const fixture = `import sys
sys.path.insert(0, sys.argv[1])
from test_regressions import MAIN, custom_body, modify
modify(sys.argv[2], 'university.docx', {MAIN: custom_body})
`
  command(python, ['-B', '-X', 'utf8', '-c', fixture, testsDir, join(base, 'resources', 'bio-thesis-master.docx')])
  const digest = file => createHash('sha256').update(readFileSync(file)).digest('hex')
  const before = digest('university.docx')
  const cli = (...args) => [ '-B', '-X', 'utf8', join(base, 'scripts', 'bio_thesis.py'), ...args ]
  command(python, cli('init', '--template', 'university.docx', '--output', 'blind.docx',
    '--title-zh', 'Example & title', '--author', 'Synthetic Author', '--review-stage', 'blind'))
  assert.equal(digest('university.docx'), before)
  const inspection = `import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile
sys.path.insert(0, sys.argv[1])
from test_regressions import MAIN, W, modify, visible
with ZipFile('university.docx') as src, ZipFile('blind.docx') as dst:
    assert src.namelist() == dst.namelist()
    assert [p for p in src.namelist() if src.read(p) != dst.read(p)] == [MAIN]
    text = visible(ET.fromstring(dst.read(MAIN)))
    assert 'Prefix Example & title and Example & title' in text and 'Synthetic Author' not in text
def change_top_margin(root):
    margin = root.find('.//' + W + 'pgMar')
    assert margin.get(W + 'top') == '2281'
    margin.set(W + 'top', '900')
modify('blind.docx', 'wrong-margin.docx', {MAIN: change_top_margin})
`
  command(python, ['-B', '-X', 'utf8', '-c', inspection, testsDir])
  const audit = (doc, report) => cli('audit', '--template', 'university.docx', '--docx', doc,
    '--review-stage', 'blind', '--sensitive-names', 'Synthetic Author', '--json-report', report)
  command(python, audit('blind.docx', 'good.json'))
  assert.equal(outcome('good.json', 'Layout-1'), 'PASS')
  assert.equal(outcome('good.json', 'Acad-10-Blind'), 'PASS')
  command(python, audit('wrong-margin.docx', 'bad.json'), 1)
  assert.equal(outcome('bad.json', 'Layout-1'), 'FAIL')
  passed = true
} finally {
  if (shutdown) await shutdown.shutdown(passed ? 0 : 1)
  process.chdir(initialDirectory)
  if (passed && !process.env.DSH_E2E_PARENT && dirname(realpathSync(work)) === parent) {
    rmSync(work, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
  } else {
    console.log(`DSH_E2E_WORK=${work}`)
  }
}
console.log('DSH_E2E=PASS install, profile boot, skill list/get, template fill, blind audit, bad margin rejected')
