# Thesis Formatting Done Right: Template-Aware & Template-Free Word Drafting + 65-Item Blind Audit

[中文说明](README.zh.md) · [No-Template Paired Benchmark](docs/no-template-comparison.md) · [3-School Template Trial](docs/intake.md) · [Head-to-Head Audit Comparison](docs/comparison.md) · [Word Evaluation](docs/evaluation.md)

**In plain English:** Whether your university hands you an official `.dotx`/`.docx` template **or only a written list of formatting rules**, this zero-dependency DeepSeek Harness (DSH) plugin handles both sides of thesis formatting:
1. **Generate clean, continue-editing-safe Word drafts**: Preserve an official school template verbatim **or adapt the built-in desensitized academic master via JSON `--profile-config`** in one command — binding every font, spacing, indentation, three-line table, and figure-caption rule directly into **19 dedicated Word top-ribbon Styles** plus ready-to-continue modular examples so human post-editing never breaks layout (instead of surface text tricks that collapse on follow-up edits);
2. **Catch silent regressions before submission**: Audit margins per section, `core.xml`/split-run blind-review leaks, italic species/genes, restriction enzyme spacing (`EcoR I` vs `EcoRI`), `µL` vs `uL`, vertical table borders, and uncited/malformed references.

---

## Measured Input -> Output Showcase (No-Template Written-Spec Task)

Given only generic written formatting rules (no `.dotx`/`.docx` file provided), rendered natively in Microsoft Word 16 — `bio-thesis` generates 15 publication-grade pages with 19 dedicated top-ribbon Styles and 0 audit failures in `<2s` (`0` custom lines), compared against bare LLM code generation (`1,212` lines -> `7` audit failures / `29` unstructured generic styles, and `gpt-6.1-sol` `183–208`-line drafts showing unstyled blue headings and missing indents):

![Word render comparison between bare coding agent and bio-thesis built-in desensitized master](docs/images/no-template-showcase.png)

| Dimension | Bare Coding Agent (`gpt-6.1-sol`) | Existing Open-Source Tools | With `dsh-plugin-bio-thesis` (`v0.3.0`) |
|---|---|---|---|
| **No-template written-spec generation** | Authored **183–208 lines** of `python-docx`; passed XML checks (`15/15`) yet rendered in Word with plain text cover, **default blue headings/captions**, unbolded TOC1, missing header rule, no 2-char indent, extra blank page | Pandoc `--reference-doc` drops template cover/frontmatter (`5/12` in multi-school trial) | **0 lines custom code**, 1 CLI command adapting the built-in master: **15/15 structural + publication-grade Word visuals** across 15 rendered pages |
| **3-school `.dotx`/`.docx` template intake** | Wrote **33–37 lines** of raw DOCX conversion code per session | `docxtpl` renders placeholders (`12/12`) without post-revision auditing | **0 lines conversion code**; all 18 outputs passed and opened in Word |
| **Pre-submission defect detection** | Requires writing ad-hoc checks each run | `format-thesis-docx` excels at auto-reformatting drafts, but caught **4/14** defects on shared rules (`DocFormatTool` **1/14**) | Native CLI caught **14/14** defects (with JSON style rules) and **8/8** planted no-template defects |
| **Blind-review leak & human checklist** | Easy to miss `docProps/core.xml` or split XML runs | Most formatters lack blind-leak XML scans + human visual checklist | Screens body, headers, rels, and properties; exports editable 65-item checklist |

*Honest engineering note:* In autonomous skill trials where the agent ran `bio_thesis.py init` and *then* wrote manual raw XML post-patches (`complete_document.py`), one session corrupted `styles.xml`/`sectPr` so Word refused to open the file. **Always drive `init --profile-config` directly via JSON/CLI flags rather than hand-patching generated OOXML.** See [No-Template Paired Benchmark](docs/no-template-comparison.md).

---

## 1-Minute Setup & Prompt

Requires Python 3.10+ (**Python standard library only — zero `pip install` needed**) and DSH `0.2.2-rc.5+`.

```sh
dsh plugin --profile web add "github:Spark-xY6/dsh-plugin-bio-thesis#v0.3.0"
```

After restarting the DSH profile, ask:

- **No template file (written rules only):**
  > Use bio-thesis-formatter to adapt the built-in master using a `--profile-config` JSON for my written formatting rules (do not hand-patch generated XML afterward) to create `draft.docx`, and audit `submitted.docx` against those rules into `review.md`.
- **With school `.dotx`/`.docx` template:**
  > Use bio-thesis-formatter to create `draft.docx` from `school.dotx` in blind review stage with title `"Example Study"`. Compare `revised.docx` with `school.dotx` and output `review.md` listing remaining Word visual checks.

### Standalone CLI

```sh
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py init --profile-config presets/custom-university.example.json --mode phd-thesis --review-stage blind --output draft.docx --title-zh "Example Study" --author "Example Author"
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py audit --docx revised.docx --profile-config presets/custom-university.example.json --review-stage blind --sensitive-names "Example Author,Example Advisor" --md-report review.md
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py export-checklist --output checklist.docx
```

---

## Benchmarks & License

- 📊 **[No-Template Comparison](docs/no-template-comparison.md)** · **[3-School Template Intake](docs/intake.md)** · **[14-Defect Audit Comparison](docs/comparison.md)**
- 🔍 **[Alternatives & Sources](docs/alternatives.md)** · **[DSH Host & Word Acceptance](docs/evaluation.md)** · **[GitHub Actions](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/actions)**
- Released under the [MIT License](LICENSE). Placeholder text/tables are structural demonstrations only; refresh TOC fields (`F9`) and inspect signatures/images in Microsoft Word before submission.
