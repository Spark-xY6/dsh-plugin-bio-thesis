# Word Thesis Layout Repair & Format Checker (`dsh-plugin-bio-thesis`)

[简体中文](README.md) · English · Pure Python stdlib (zero third-party installs) · Read-only on original files

**Start a clean draft with all 19 Word Ribbon styles and an Author CV section locked in under 2 seconds when you only have written rules—or compare an existing multi-section thesis against your school `.dotx`/`.docx` template, preview exact centimetre and numbering diffs per section, save a repaired copy with landscape wide tables intact, and run a 67-item multi-category pre-submission audit.**

![Locked Word Ribbon styles, triple-section isolation, and 67-item multi-category audit showcase](docs/images/no-template-showcase.png)

## What to ask your assistant

| Goal | Prompt |
| --- | --- |
| Align an existing thesis | “Compare draft.docx against school.dotx with bio-thesis-formatter, preview margin and page-number diffs per section, then save draft-repaired.docx.” |
| Start from written rules | “Use bio-thesis-formatter to generate thesis.docx matching these formatting requirements, with headings, body, captions, three-line tables, and Author CV locked into Word styles.” |
| Rigorous pre-submission audit | “Audit draft.docx against school.dotx for layout, figure font size (>= 7.5 pt), third-person tone, online-only DOIs, unified time units, Chinese punctuation spacing, Author CV, and blind-review redaction.” |

## Why it exists, what it does & verified results

- **No fragile 1,000-line ad-hoc scripts**: Bare LLM scripts routinely pile unbound styles into Word that collapse once you start editing (`1,212` lines of bare Python still left `7` formatting failures in our benchmark). This tool compiles rules cleanly into **19 live Word Ribbon styles** (switching seamlessly from e.g. 20 pt Songti/Nature to 22 pt Fangsong/GB-T 7714) and builds a **15-page structured draft (with Author CV in `final` mode and automatic identity redaction in `blind` mode) in under 2 seconds** with `0` ad-hoc code.
- **Role-aware multi-section template repair**: Maps front matter, Roman-numbered abstracts, Arabic main chapters, and landscape wide-table islands cleanly onto 3-section school `.dotx`/`.docx` templates—syncing paper size, all 7 `pgMar` fields, mirror margins, even/odd headers, and chapter page-1 restarts while keeping wide tables landscape.
- **Rigorous default 67-item multi-category checks**: Whenever no custom override is provided, built-in checks automatically enforce strict academic details: **figure text `>= 7.5 pt` (`6 号字`)**, **no first-person prose (`我`/`我们`/`本人`) in main chapters**, **DOIs permitted only on ahead-of-print / preprint references without formal volume/page numbers**, **unified Chinese or English experimental time units**, **Chinese commas inside multi-panel citations (`图 1.1 A，B`) and equipment parentheticals (`P-2000，Sutter Instrument`)**, **script-matched spaces (half-width around English/numerals, full-width between Chinese items)**, and **Author CV completeness / blind-review redaction**. Catches **`14/14` planted structural & identity defects with `0` false positives** across 8 test drafts (vs `1/14` and `4/14` for comparable open-source tools).

## Install & CLI

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-bio-thesis-0.5.0.tgz').Path
```

```powershell
$cli = 'skills/bio-thesis-formatter/scripts/bio_thesis.py'
python -B $cli repair --docx draft.docx --template school.dotx
python -B $cli repair --docx draft.docx --template school.dotx --output draft-repaired.docx
python -B $cli audit --docx draft-repaired.docx --template school.dotx --base-docx draft.docx --md-report review.md
```

[Template trial](docs/intake.md) · [Benchmark comparison](docs/comparison.md) · [Written rules trial](docs/no-template-comparison.md) · [Evaluation](docs/evaluation.md) · [MIT License](LICENSE)
