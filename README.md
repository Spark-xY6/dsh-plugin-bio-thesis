# Thesis Formatting Without Fragile Hacks · bio-thesis

[中文说明](README.zh.md) · [No-Template Benchmark](docs/no-template-comparison.md) · [3-School Template Trial](docs/intake.md) · [14-Defect Audit](docs/comparison.md) · [Word Evaluation](docs/evaluation.md)

Whether your department gives you an official `.dotx`/`.docx` file **or only a plain-text list of formatting rules**, letting a coding agent piece together raw Word XML from scratch leaves styles unbound—so the first time you edit a paragraph or insert a table in Microsoft Word, headings, indents, and spacing fall apart.

Using Python's standard library only (zero third-party `pip` packages), `dsh-plugin-bio-thesis` preserves official school templates cleanly **or adapts its built-in desensitized master via `--profile-config` JSON**—binding every heading, body paragraph, three-line table, figure caption, table title, and bibliography rule into **19 dedicated `qFormat` Word top-ribbon Styles** alongside ready-to-edit examples (`图 1.1` / `表 1.1`). Before submission, the audit tool catches shifted section margins, `core.xml` and split-run blind-review leaks, italic genes/species, `µL` typos, vertical table lines, and uncited references.

![Word top-ribbon 19 locked styles, redacted cover, isolated TOC, and three-line table render](docs/images/no-template-showcase.png)

## Measured Benchmarks & Honest Scope

- **Writing from written rules only (no template file)**: On a neutral formatting spec, an unassisted coding subagent wrote `1,212` lines of Python yet tripped `7` audit failures (`29` unstructured styles, missing caption/bibliography styles, left-shifted cover date, and an out-of-range `[9]` reference). With `bio-thesis`, one command creates a `FAIL: 0` 15-page draft in under `2s` (`0` custom lines), and switching `--profile-config` JSON cleanly adapts `word/styles.xml` across universities (`某高校 SimSun 20pt / Nature` to `乙高校 FangSong 22pt / GB/T 7714`, both `FAIL: 0`).
- **3-school `.dotx`/`.docx` intake & open-source audit comparison**: Zero glue code written across `18/18` Word-opened deliverables (`33–37` ad-hoc lines when bare). Across shared 7-school rules with `14` planted defects, the native auditor caught **`14/14` (100%)**, compared against `format-thesis-docx` (`4/14`) and `DocFormatTool` (`1/14`). Full logs are in [docs/no-template-comparison.md](docs/no-template-comparison.md) and [docs/comparison.md](docs/comparison.md).
- **Important guard**: Do not let an agent manually patch raw OOXML after running `bio_thesis.py init`—in one earlier autonomous session, manual post-edit scripts corrupted `styles.xml`/`sectPr` so Word refused to open the file. Always configure fonts, margins, and citation styles through `--profile-config` JSON and CLI flags.

## 1-Minute Install & CLI

Requires Python 3.10+ (standard library only) and DeepSeek Harness `0.2.2-rc.5+`:

```powershell
dsh plugin --profile web add "github:Spark-xY6/dsh-plugin-bio-thesis#v0.3.0"
```

After restarting the profile, send whichever prompt fits your files:

> **Written rules only (no template):** Use bio-thesis-formatter to adapt the built-in master via `--profile-config` JSON for my formatting rules (do not hand-patch generated XML afterward) to create `draft.docx`, and audit `submitted.docx` against those rules into `review.md`.

> **Official `.dotx`/`.docx` template on hand:** Use bio-thesis-formatter to create `draft.docx` from `school.dotx` in blind review stage with title `"Example Study"`. Compare `revised.docx` with `school.dotx` and output `review.md` listing visual items left to verify in Word.

Or run offline without DSH:

```sh
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py init --profile-config presets/custom-university.example.json --mode phd-thesis --review-stage blind --output draft.docx --title-zh "Example Study" --author "Example Author"
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py audit --docx revised.docx --profile-config presets/custom-university.example.json --review-stage blind --sensitive-names "Example Author,Example Advisor" --md-report review.md
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py export-checklist --output checklist.docx
```

Shipped placeholder tables and citations are layout demonstrations only; refresh TOC page numbers in Word (`F9`) and inspect signatures and embedded images before submission. Additional references: [docs/intake.md](docs/intake.md), [docs/alternatives.md](docs/alternatives.md), [docs/evaluation.md](docs/evaluation.md). [MIT License](LICENSE).
