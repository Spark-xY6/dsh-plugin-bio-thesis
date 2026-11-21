---
name: bio-thesis-formatter
description: Create Word thesis drafts from school DOCX or DOTX templates, audit existing drafts for structural and formatting issues, screen specified blind-review identities, and export a 65-item review checklist.
whenToUse: Use when a student creates a Word thesis from an institutional template, checks a revised thesis or milestone report, or prepares a blind-review copy.
disable-model-invocation: false
user-invocable: true
metadata:
  version: "0.3.0"
  license: MIT
---

# Bio-Thesis Formatter

Create a Word thesis or milestone report from the student's template and audit later drafts against it. Apply discipline-specific typography and citation checks when requested.

Resolve all relative paths below (`scripts/`, `resources/`, `references/`) against this skill's base directory (`<skill_resources>` in DeepSeek Harness).

## Scope & Tool Separation

Start with the user's template and chosen rules. Use the CLI to create a new draft or audit an existing one. The editable checklist includes automatic checks, visual inspection and human judgment. Cite individual results and their scope when reporting to the author.

## Python & Runtime Environment

1. Follow any explicit user or project `AGENTS.md` Python environment instruction first.
2. In DeepSeek Harness environments that expose `load_workspace_dependencies`, call it to locate the bundled Python executable; otherwise run directly with `python3` or `python` (`>= 3.10`) on `PATH`.
3. Every helper script in `scripts/` uses **Python standard library only** (`xml.etree.ElementTree`, `zipfile`, `json`, `argparse`, `pathlib`) — **no `pip install` or third-party dependencies required**. Avoid creating `__pycache__` inside the skill folder by running Python with `-B` (`python3 -B ...`).
4. When visual layout inspection (page breaks, blank verso pages, refreshed TOC rendering) is requested inside DeepSeek Harness and `office_render` / `office_convert` tools exist (via the `office` skill), use them on a workspace copy; otherwise instruct the user to update the TOC field (`F9`) in Word and verify print PDF export visually.

## Safety & Input Integrity

- **Never overwrite an author's original `.docx` without explicit permission.** Generate new files or formatted candidates alongside the baseline and pass `--base-docx <original.docx>` during audit so embedded `word/media/*` image hashes are verified intact (`Pkg-Media`).
- **Keep outputs in the user's workspace**, never inside the skill directory itself (`audit_thesis_docx.py` refuses writing reports inside the skill directory).
- Distinguish automated **OOXML XML checks** from **visual Word/PDF pagination checks** — never claim a document passed visual pagination inspection unless pages were actually rendered and inspected.

## Standard 4-Step Workflow

### Step 1: Select Preset, Document Mode & Review Stage

Review [document-modes-and-presets.md](references/document-modes-and-presets.md):

- **Presets (`--preset`)**:
  - `cn-lifesci-default` (Nature superscript citations, standard Chinese biomedical A4 geometry)
  - `cn-gb7714-general` (GB/T 7714-2015 `[1]` sequential citations)
  - Or pass `--profile-config <custom.json>` / `--university` / `--school` / `--citation-style` to match any institution.
- **Modes (`--mode`)**: `phd-thesis`, `master-thesis`, `proposal`, `midterm`, `rotation`.
- **Stages (`--review-stage`)**: `final` (full author/advisor/acknowledgements) vs `blind` (redacts cover author/ID/advisor, English abstract names, CV authors, acknowledgements, and marks collaborative caption phrasing anonymously).

### Step 2: Initialize Clean `.docx` OR Inspect Existing Draft

Prefer the institution's actual Word template when the user supplies one. Pass `init --template <school.docx|school.dotx> --output <candidate.docx>`. Properties, sections, styles, fields, media and content come from that template. Fill explicit markers in ordinary Word text, such as `{{title_zh}}` with `--title-zh` or `{{author}}` with `--author`. Template mode masks identity markers in blind mode. Check actual identity-bearing text, the file properties, images, signatures and embedded files before sending.

Pass `audit --template <school.docx|school.dotx>` to use that template as the layout reference. Supply the applicable citation style and use `audit.template_checks` in the optional JSON profile to opt into other school-specific checks. Treat `INFO` coverage gaps and `WARN` findings as review items. Generic settings below apply to the bundled master; inspect the supplied template and local requirements for other institutions.

If the school requires a named paragraph style to control font size, family or colour, put the corresponding rule in a JSON profile and pass it with `audit --profile-config <rules.json>`. Read the exact paragraph `w:styleId` from the template's `word/styles.xml`. Replace the example ID and select only properties covered by the school's rule.

```json
{"audit":{"direct_formatting":[{"id":"body-size","paragraph_style":"SchoolBody","forbid":["font_size"]}]}}
```

`font_size` covers regular and complex-script size, `font_face` covers run font assignments, and `color` covers text colour. The check flags these explicit assignments even when they currently equal the inherited value. Named character styles, inline bold and italics are allowed. Each rule covers its paragraph style and its descendants, including the implicit default style where applicable. A finding such as `DirectFormat-body-size` gives the one-based paragraph position in main-document order, including table cells. This check inspects current text runs and skips deleted text, text boxes, headers, footers and notes. An unrecognised rule or style ID exits with a configuration error. A valid style with no matching paragraphs produces a coverage warning.

Scaffold a clean `.docx` (optionally embedding a custom university PNG/JPEG emblem via `--logo`) using `scripts/bio_thesis.py`:

```sh
python3 -B scripts/bio_thesis.py init \
  --output thesis_draft.docx \
  --preset cn-lifesci-default \
  --mode phd-thesis \
  --review-stage blind \
  --university "Your University" \
  --school "School of Life Sciences" \
  --title-zh "中文题目" \
  --title-en "English Title" \
  --author "Author Name" \
  --advisor "Advisor Name"
```

Export an editable 65-item checklist `.docx` at any time:

```sh
python3 -B scripts/bio_thesis.py export-checklist --output checklist_65.docx
```

### Step 3: Apply the selected institution's rules

The following conventions describe the bundled example. Apply them to external Word templates only if the institution requests them. See [layout](references/ooxml-layout-and-styles.md) and [typography](references/typography-tables-figures-citations.md).

- **Three-Section Isolation**: `S1` Cover + Copyright (no header, no page number) -> `S2` Abstracts + TOC (`upperRoman` from `I`) -> `S3` Main Chapters + References + CV + Acknowledgements + Declaration (`decimal` from `1`), with `mirrorMargins` & `evenAndOddHeaders`.
- **Three-Line Tables**: Top/bottom borders `1.5 pt` (`sz="12"`), header bottom border `0.75 pt` (`sz="6"`), zero vertical/interior lines (`nil`), `tblHeader` + `cantSplit`.
- **Bilingual Fonts & Biological Nomenclature**: Chinese in SimSun/SimHei; Latin, numbers, Greek (`α`, `μ`), and `°C` in `Times New Roman`; genes & Latin binomials italicized (`TP53`, *Mus musculus*); restriction enzymes italicized with a space before Roman numerals (*EcoR I*, *Taq I*); `µL` (never `uL`), space before units except `%` and `°C`; consistent Chapter 2 time units; clean dashes (`-`, `–`, `—`, `~`).
- **Nature vs GB/T 7714 Citations**: Nature mode checks superscript placement before punctuation, spaced initials, `&`, sentence-case titles, italicized abbreviated journals, bold volume, en-dash pages, `(Year)` without trailing period, English refs before Chinese refs, and zero residual `EN.CITE` codes. GB/T 7714 mode checks bracketed `[1]` citations, document type tags (`[J]`/`[M]`/`[D]`), and bidirectional reference closure.

### Step 4: Audit the candidate and inspect the rendered document

```sh
python3 -B scripts/bio_thesis.py audit \
  --docx candidate.docx \
  --base-docx original.docx \
  --mode auto \
  --review-stage final \
  --preset cn-lifesci-default \
  --genes "Tp53,Brca1" \
  --species "Mus musculus,Homo sapiens" \
  --json-report audit.json \
  --md-report audit.md
```

Map every `FAIL`/`WARN` finding to [self-check-65-matrix.md](references/self-check-65-matrix.md), resolve structural/typography issues in a new output copy, then finish with manual Word `F9` TOC update and print/PDF visual inspection.

Use `audit --fail-on-warn` for an opt-in strict submission gate. Include pending tasks from `manual_review` when summarizing the JSON report. Full reports may contain document text and local paths.
