---
name: bio-thesis-formatter
description: Create Word thesis drafts from school DOCX or DOTX templates, audit an existing draft against templates or rigorous built-in 67-item life-science dissertation defaults, preview or repair section margins, screen blind-review identities, and export the 67-item checklist.
whenToUse: Use when a student creates a Word thesis from an institutional template, checks a revised thesis or milestone report, or prepares a blind-review copy.
disable-model-invocation: false
user-invocable: true
metadata:
  version: "0.5.0"
  license: MIT
---

# Bio-Thesis Formatter

Create a Word thesis or milestone report from the student's template and audit later drafts against it. Apply discipline-specific typography and citation checks when requested.

Resolve all relative paths below (`scripts/`, `resources/`, `references/`) against this skill's base directory (`<skill_resources>` in DeepSeek Harness).

For an existing manuscript, route “check my draft” to `audit`. For a full role-aligned layout comparison (section geometry, all margins/headers/footers/gutter, `mirrorMargins`/`evenAndOddHeaders`, and section page numbering), run `scripts/bio_thesis.py repair --docx DRAFT --template SCHOOL`; add `--output NEW.docx` to save a repaired copy alongside optional `--no-fix-numbering`, `--fix-tables`, or `--clean-leading-spaces` flags (keep `repair-margins` when only previewing or syncing the four margins). Match sections by role (cover, front-matter, body, landscape wide-table islands). After a repair, use `audit --docx NEW.docx --template SCHOOL --base-docx DRAFT` to distinguish remaining issues. Use `init` when the user wants a fresh draft.

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

Support two equally first-class paths depending on what the author supplies:

- **Path A — No school template file (written formatting requirements only)**: Use the bundled desensitized academic master (`resources/bio-thesis-master.docx`) adapted with `--preset` + `--profile-config <custom.json>` (overriding page dimensions, margins in cm, `mirror_margins`, `odd_even_headers`, `typography`, `university_name`, `school_name`, `citation_style`, `modes`, `audit.disabled_checks`, `audit.forbid_published_doi`, `--mode`, and `--review-stage`). **NEVER write ad-hoc Python/OOXML scripts to hand-patch `word/styles.xml` or `sectPr` tags on top of the `.docx` produced by `bio_thesis.py init`** — hand-patching generated OOXML risks Word corruption/open errors; pass parameters or `--profile-config` JSON directly to `bio_thesis.py init` instead.
- **Path B — School `.docx` / `.dotx` template supplied**: Pass `init --template <school.docx|school.dotx> --output <candidate.docx>`. Properties, sections, styles, fields, media and content come from that template. Fill explicit markers in ordinary Word text, such as `{{title_zh}}` with `--title-zh` or `{{author}}` with `--author`. Template mode masks identity markers in blind mode. Check actual identity-bearing text, the file properties, images, signatures and embedded files before sending.

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

Export an editable 67-item checklist `.docx` at any time:

```sh
python3 -B scripts/bio_thesis.py export-checklist --output checklist_67.docx
```

### Step 3: Apply the selected institution's rules (Built-In Rigorous Defaults When Unspecified)

Unless the author provides an external `--template` or custom overrides via `--profile-config`, enforce these built-in default dissertation rules. See [layout](references/ooxml-layout-and-styles.md) and [typography](references/typography-tables-figures-citations.md).

- **Three-Section Isolation & Author CV (`Acad-CV`)**: `S1` Cover + Copyright (no header, no page number) -> `S2` Abstracts + TOC (`upperRoman` from `I`) -> `S3` Main Chapters + References + Author CV & Publications (`个人简历及在学期间发表的学术论文与研究成果`, redacted in `--review-stage blind`, required by default in `final` degree/proposal/midterm modes) + Acknowledgements + Declaration (`decimal` from `1`), with `mirrorMargins` & `evenAndOddHeaders`.
- **Three-Line Tables & Figure Standards (`Fig-5`)**: Top/bottom borders `1.5 pt` (`sz="12"`), header bottom border `0.75 pt` (`sz="6"`), zero vertical/interior lines (`nil`), `tblHeader` + `cantSplit`. In-figure labels must be `>= 6 号字` (`7.5 pt`) by default (`typography.figure_min_font_pt`), checked automatically on DrawingML/SVG runs (`Fig-5`) and verified visually on exported PDFs.
- **Academic Prose, Punctuation, Spacing & Nomenclature (`Acad-14`, `Err-2-11-12`, `Err-7`, `Err-14`)**: No first-person pronouns (`我`/`我们`/`本人`) outside legal declarations; consistent objective self-term (`本论文`/`本研究`/`本工作`). Experimental procedure time units must stay uniformly Chinese (`小时`/`分钟`/`秒`) or English (`h/hr`, `min`, `s/sec`), never mixed. Equipment/vendor pairs (`P-2000，Sutter Instrument`) and multi-subpanel references (`图 1.1 A，B`) use Chinese commas `，`. Use half-width ASCII spaces after English/numerals and Chinese full-width spaces (`　`) when separating Chinese items (`姓名：XXX　性别：X`). Genes & Latin binomials italicized; restriction enzymes italicized with a space before Roman numerals (*EcoR I*, *Taq I*); `µL` (never `uL`); clean dashes (`-`, `–`, `—`, `~`).
- **Nature vs GB/T 7714 Citations (`Ref-2-7`)**: Nature mode (default) checks superscript placement before punctuation, spaced initials, `&`, sentence-case titles, italicized abbreviated journals, bold volume, en-dash pages, `(Year)` without trailing period, English refs before Chinese refs, zero residual `EN.CITE` codes, and **no DOIs except on online-ahead-of-print/preprint references lacking volume/pages** (`audit.forbid_published_doi`). GB/T 7714 mode checks bracketed `[1]` citations, document type tags (`[J]`/`[M]`/`[D]`), and bidirectional reference closure.

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

Map every `FAIL`/`WARN` finding to [self-check-67-matrix.md](references/self-check-67-matrix.md), resolve structural/typography issues in a new output copy, then finish with manual Word `F9` TOC update and print/PDF visual inspection.

Use `audit --fail-on-warn` for an opt-in strict submission gate. Include pending tasks from `manual_review` when summarizing the JSON report. Full reports may contain document text and local paths.
