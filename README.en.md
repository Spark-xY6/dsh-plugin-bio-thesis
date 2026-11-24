# Word Thesis Layout Repair & Format Checker (`dsh-plugin-bio-thesis`)

[简体中文](README.md) · English · Pure Python stdlib (zero extra installs) · Original files stay read-only

Lock **19 Word Ribbon styles** and an Author CV into clean thesis drafts in under 2 seconds, align existing multi-section manuscripts against your school `.dotx` template without breaking landscape tables, and run a **67-item pre-submission & blind-review check**—always saving onto a separate editable copy.

![Locked Word Ribbon styles, triple-section isolation, and 67-item multi-category audit showcase](docs/images/no-template-showcase.png)

## When it helps

- **Starting with written rules and no template**: Generates a 15-page structured draft (headings, body, captions, three-line tables, F9 TOC, and Author CV) locked into **19 live Word styles** in `< 2` seconds, with instant preset switching and automatic blind-review redaction (vs `1,212` lines of bare script still leaving `7` Word defects).
- **Aligning a multi-section thesis with wide tables**: Maps front matter, Roman abstracts, Arabic chapters, and landscape wide-table islands cleanly onto `.dotx`/`.docx` school templates—previewing exact centimetre & page-number diffs before saving a repaired copy.
- **Catching subtle typography & identity leaks**: Built-in **67-item checks** verify figure font size (`>= 7.5 pt` / `6 号字`), objective third-person prose, online-only DOIs, unified experimental time units, Chinese commas inside subpanels & instrument parentheticals, script-matched spacing, and zero author/metadata leaks (`14/14` planted pitfalls caught across 8 test drafts).

## How to use

1. Install into DSH:

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-bio-thesis-0.5.0.tgz').Path
```

2. Talk to your assistant (or run `skills/bio-thesis-formatter/scripts/bio_thesis.py` directly via CLI):

| Goal | Prompt |
| --- | --- |
| Align an existing thesis | “Compare draft.docx against school.dotx with bio-thesis-formatter, preview margin and page diffs, then save draft-repaired.docx.” |
| Start from written rules | “Use bio-thesis-formatter to generate thesis.docx matching these rules, with styles and Author CV locked in Word.” |
| Pre-submission & blind audit | “Audit draft.docx against school.dotx for layout, figure font size (>= 7.5 pt), third-person tone, online-only DOIs, spacing, and blind-review leaks.” |

[Template trial](docs/intake.md) · [Benchmarks](docs/comparison.md) · [Written rules trial](docs/no-template-comparison.md) · [Evaluation](docs/evaluation.md) · [MIT License](LICENSE)
