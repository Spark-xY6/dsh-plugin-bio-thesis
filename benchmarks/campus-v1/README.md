# Reproduce the Campus Word benchmark

All documents, identities and artwork are synthetic. The fixture builder uses Node's standard library. The scorer and the native comparison adapter use Python's standard library. Work in an empty scratch directory for outputs. Use Node 24 and Python 3.10+ for the benchmark. For Windows with conda, run the Python commands in your activated `base` environment. Keep the evaluator, gold labels and other contestants' output outside every model workspace.

## Rebuild and check the corpus

From the repository root run the following. `build.mjs` creates ten documents under `benchmarks/campus-v1/inputs` and refuses to replace mismatched inputs. Git ignores the generated Word files.

```sh
node benchmarks/campus-v1/build.mjs
python -B -X utf8 benchmarks/campus-v1/score.py --self-check
```

An independent OOXML reader must match the published gold labels for all eight files. Eleven inputs, including the assignment, have recorded SHA-256 values. The complete manifest SHA-256 is `371b464fbfb883e24235c7d4e0b3fa35af6a94738af921179e50afc0ccfbb44b`. [Assignment](inputs/TASK.md) · [Predeclared decisions](design.json) · [Review and frozen hashes](freeze.md)

## Reproduce the primary native comparison

The following assumes a fresh sibling directory named `campus-repro`. It uses the exact plugin revision evaluated in the model comparison and an unmodified external comparator. Keep this pinned checkout even when running from newer plugin versions.

```sh
git clone https://github.com/Spark-xY6/dsh-plugin-bio-thesis ../campus-repro/plugin
git -C ../campus-repro/plugin checkout --detach 4547e792ef7722b4056fdce57d0df4223b513084
git clone https://github.com/wksudud/format-thesis-docx ../campus-repro/format-thesis-docx
git -C ../campus-repro/format-thesis-docx checkout --detach c04f786d5b994879f938325630e9ffdf4dde00be
python -B -X utf8 benchmarks/campus-v1/compare_cli.py --inputs benchmarks/campus-v1/inputs --plugin ../campus-repro/plugin --upstream ../campus-repro/format-thesis-docx --out ../campus-repro/native
python -B -X utf8 benchmarks/campus-v1/score.py --output ../campus-repro/native/v020
python -B -X utf8 benchmarks/campus-v1/score.py --output ../campus-repro/native/format-thesis-docx
```

The adapter translates native reports, never the gold labels. It accepts a named template-style size warning as STYLE. For the MIT tool it reads the extracted requirements and native CSV, counting explicit failures in the measured properties. Manual rows have no automatic label. A missing generation or audit is assigned zero by the joint-task model scorer. For specialist tools the public comparison records only the applicable subtask and marks the other as null. [Native results and per-document labels](native-results.json)

To repeat the post-hoc v0.3.0 regression, run the following from this release's root, changing A05 to each ID A01 through A08. The opt-in profile bans direct point-size formatting within the school body style. In addition to the original adapter's labels, count DIRECT when the report contains `DirectFormat-body-size` with status `FAIL`. For the two positive files the native detail is `paragraph=7 properties=font_size`. Perform the unchanged `init` step from the adapter and score the two outputs. The machine-readable snapshot records the exact profile and cases.

```sh
python -B -X utf8 skills/bio-thesis-formatter/scripts/bio_thesis.py audit --docx benchmarks/campus-v1/inputs/A05.docx --template benchmarks/campus-v1/inputs/school.docx --base-docx benchmarks/campus-v1/inputs/school.docx --review-stage blind --sensitive-names EXAMPLE-Q17 --profile-config benchmarks/campus-v1/campus-profile.json --json-report ../campus-repro/A05-v030.json
```

To repeat the MIT tool's own repair example, first extract requirements from its `examples/template.docx` with `scripts/extract_requirements.py --template TEMPLATE --out REQS`. Run `scripts/audit_format.py --requirements REQS/requirements.json --paper PAPER --out FIRST` with `PAPER=examples/paper_before.docx`. Then run `scripts/apply_format.py --requirements REQS/requirements.json --audit FIRST/format_audit.csv --paper PAPER --out FIXED`. Re-run the audit on `FIXED/paper_before_formatted.docx` and compare the two native CSV reports.

## Same-model procedure

Use Codex CLI 0.169.0 on Windows and a model account with access to `gpt-6.1-sol`. For each of three pairs, start two separate ephemeral sessions with a fresh parent working directory, identical `inputs/` and an empty `output/`. Copy **only** the eleven manifest entries into each `inputs/`. For the treatment session install the 13 files of `skills/bio-thesis-formatter/` from the pinned plugin into `.agents/skills/bio-thesis-formatter/`. The bare session gets an empty skill inventory. Verify the treatment against the v0.2.0 release tarball, SHA-256 `818d51c8294a3875cac8884cd02f484060302d8297d7be3734e2f56bb639ecc9`.

Use this identical prompt in both conditions, with the working directory set to that session's isolated directory:

> Complete both tasks described in inputs/TASK.md. Use the seven labels and all eight drafts. Place result.json and generated.docx in output/. Use the provided local skill and its scripts if this workspace contains any under .agents/skills; supplement with your own local checks if needed. Otherwise use ordinary local tools and code. A Windows Python interpreter is available on PATH. Work only inside the current workspace. Do not open parent/sibling directories, other projects, past conversations, any user-level skills or the network. Leave all files under inputs/ intact. Finish when the two requested output files exist.

The actual settings were `exec --ignore-user-config --ephemeral --skip-git-repo-check --enable skip_host_skill_discovery --disable skill_search --disable plugins --disable apps --disable memories --disable multi_agent -c skills.bundled.enabled=false -c skills.include_instructions=false -c project_doc_max_bytes=0 -c suppress_unstable_features_warning=true -c model_reasoning_effort="low" -c web_search="disabled" -c sandbox_workspace_write.network_access=false -c windows.sandbox="elevated" -m gpt-6.1-sol -s workspace-write --json -`. Feed the prompt on stdin. Pass model credentials through the host's normal login. Run an isolated shell read/write probe before beginning. The model uses the service, while tools inside the model's workspace have networking disabled.

Set a 600-second cap per session. Start all three pairs together and run each pair serially, in orders bare/skill, skill/bare and bare/skill. Measure wall time from spawning the solver to its exit. Count completed tool events, excluding agent messages, reasoning, errors and to-do lists. Both shell invocations and file-change events count. One bare run used five shell invocations and one file-change event. Check both result files, all eleven input digests, command exit states and evidence of actual skill use. For each isolated `output/`, run `python -B -X utf8 benchmarks/campus-v1/score.py --output OUTPUT` outside the solver workspace. Retain every valid pair. The six model sessions are stochastic repeats of one fixed assignment.

## Other tools

Use sibling scratch directories for each output, create them before calling the tools, and provide their required dependencies in an appropriate environment. Use the same school DOCX for tools requiring DOCX input. The DOTX and DOCX have equivalent content and geometry and different main-part content types. Third-party code is available from its maintainers.

| Tool | Reproduction |
|---|---|
| docxtpl 0.20.2 | In the directory containing `school.docx` run `python -c "from docxtpl import DocxTemplate; d=DocxTemplate('school.docx'); d.render({'title_en':'Forms of Evidence','author':'Anonymous'}); d.save('generated.docx')"` on a scratch copy, then score its output directory. [Package](https://pypi.org/project/docxtpl/0.20.2/) |
| Pandoc 3.8 | From the root, run `pandoc -f docx+styles -t docx benchmarks/campus-v1/inputs/school.docx --reference-doc=benchmarks/campus-v1/inputs/school.docx --lua-filter=benchmarks/campus-v1/pandoc-campus.lua -o ../campus-repro/pandoc/generated.docx`, then score only generation. [Two-substitution filter](pandoc-campus.lua) |
| DocFormatTool `40f22ca37f4b856ad6d59a3382c1570e5d7ca67d` | In the checkout, run `python tools/analyze.py SCHOOL.docx -o RAW.json`, then `python tools/validate.py DRAFT.docx SCHOOL.docx --config RAW.json` for all eight drafts. Repeat with an explicit policy copy of RAW.json with `page.paper="Letter"`, `fonts.body.size_pt=13` and `page.margins_cm` set to the published four margins multiplied by 2.54/1440. Missing-run-font warnings identify no local point-size override. The native report compares image counts. [Project](https://github.com/NeoHotaru/DocFormatTool) |
| Tsinghua checker `f19b8392a111eacff7b6e80211560298abca852e` | Export each original DRAFT.docx to a PDF with Word, read-only. In the tool's checkout run `python -B -m guide_checker.cli.run_full_review --docx DRAFT.docx --pdf DRAFT.pdf --output-dir OUT` with absolute input and output paths and one output folder per draft. For the compatibility arm, from the plugin root run `python -B benchmarks/campus-v1/with-settings.py benchmarks/campus-v1/inputs ../campus-repro/with-settings`. Repeat all eight native checks using these copies and the same PDFs. Preserve the eight original incompatibilities separately. [Checker](https://github.com/jacky-rui/tsinghua-thesis-format-checker-skill) |

The public results omit absolute workstation paths, private traces and third-party manuscript examples. [Interpretation and limitations](../../docs/comparison.md)
