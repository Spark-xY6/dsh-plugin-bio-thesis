# Word thesis templates and formatting checks for DeepSeek Harness

[中文使用说明](README.zh.md) · [Existing tools and design choices](docs/alternatives.md) · [Evaluation and Word example](docs/evaluation.md) · [Downloads](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/releases)

Start an editable Word thesis from your university’s DOCX or DOTX. After revisions, compare the draft with the same template to check sections, margins and page numbering. The template supplies the cover, styles, headers, fields and pictures. This works across disciplines. Scientific typography and citation conventions are optional.

## Why this exists

A template helps a student start correctly. It cannot tell them whether a section margin changed later, an illustration was replaced, or an author name survived in file properties. This plugin brings those checks into DSH. We reviewed official school templates, Pandoc, a DOCX reformatter, a web formatting service and an open-source thesis checker. Two ideas carried into this release are an optional warning gate and an explicit list of pending human checks. [Comparison and sources](docs/alternatives.md)

## Install in DSH

Use a DSH environment with pnpm and Python 3.10+. Our host test uses DSH 0.2.2-rc.5 on Node 24. Install into the profile you use, then restart that profile.

```sh
dsh plugin --profile web add "github:Spark-xY6/dsh-plugin-bio-thesis#v0.2.0"
```

Alternatively download the prebuilt tarball from the [release](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/releases/tag/v0.2.0) and run `dsh plugin --profile web add ./dsh-plugin-bio-thesis-0.2.0.tgz`. In your DSH workspace, ask:

> Use bio-thesis-formatter to create draft.docx from school.dotx. Fill the title with “Example study”. Compare revised.docx with school.dotx and save a Markdown report listing what I should check in the final PDF.

## Standalone CLI

Run from the repository root with Python 3.10 or later. The document tools use only the Python standard library.

```sh
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py init --template school.dotx --output draft.docx --title-en "Example study"
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py audit --docx revised.docx --template school.dotx --md-report review.md --json-report review.json
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py export-checklist --output checklist.docx
```

- **Templates.** Use unsigned, macro-free DOCX or DOTX files. Plain-text markers such as `{{title_en}}` and `{{author}}` can be filled automatically. An unmarked template also yields an editable copy. Existing manuscripts can be audited. Migrating a complete manuscript to a different template requires separate editing.
- **Rules.** Add `--base-docx original.docx` to detect changes to embedded media by path and SHA-256. Select `--citation-style nature` or `gb7714` for supported numeric citation patterns. Opt into additional school rules using `audit.template_checks` in a JSON profile. Omit the template to try a built-in starter.
- **Blind review.** Use `--review-stage blind` when filling identity placeholders. Audit a completed draft with `--review-stage blind --sensitive-names "Example Author,Example Advisor"`. Inspect signatures, pictures and embedded attachments separately. Full reports can contain local paths, matched words and excerpts. The standalone CLI works locally. In chat, file access and model routing follow your DSH configuration.
- **Results.** An audit exits 1 on a blocking finding, 0 when it completes without one, and 2 for an input or execution error. `--fail-on-warn` also treats warnings as blocking. The report lists pending human checks separately. The 65-item editable checklist includes items requiring visual inspection or judgment.

## Results on synthetic documents

A draft generated from a landscape Letter template retained its header and custom title style. Changing its top margin from 1.25 to 0.5 inches and reinserting a supplied identity produced two specifically identified failures. Microsoft Word opened a separate generated draft, showing the expected header, page field, title and embedded picture. [Evidence, limitations and reproduction](docs/evaluation.md) · [Continuous integration](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/actions)

[MIT license](LICENSE)
