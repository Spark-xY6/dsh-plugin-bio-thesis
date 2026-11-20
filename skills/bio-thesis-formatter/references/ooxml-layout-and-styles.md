# OOXML Layout, Section Architecture & Semantic Style IDs

Clean-room OOXML definitions shipped in `resources/bio-thesis-master.docx` and validated by `scripts/audit_thesis_docx.py`.

## 1. Default A4 Geometry & `word/settings.xml`

Default preset values (overridable via `--profile-config <json>`):

| Property | Default OOXML Value | Physical Measure |
| :--- | :--- | :--- |
| Paper size (`pgSz`) | `w="11906" h="16838" code="9"` | A4 (`21.0 × 29.7 cm`) |
| Margins (`pgMar`) | `top="1701" bottom="1418" left="1474" right="1474"` | `3.0 / 2.5 / 2.6 / 2.6 cm` |
| Gutter / Header / Footer | `gutter="0" header="1134" footer="992"` | `0 / 2.0 / 1.75 cm` |
| Global layout flags | `<w:mirrorMargins/>`, `<w:evenAndOddHeaders/>` | Duplex symmetry & odd/even headers |
| Footnotes (`footnotePr`) | `decimalEnclosedCircleChinese`, `eachPage` | Circled numbers restarted per page |

## 2. Three-Section Header & Pagination Isolation (`Layout-SecNum`)

| Section | Included Blocks | Header Rels | Footer Rels | `<w:pgNumType>` |
| :--- | :--- | :--- | :--- | :--- |
| **Section 1** | Cover page + Copyright notice | Omitted | `footer1.xml` (empty) | Omitted (unnumbered) |
| **Section 2** | Chinese Abstract, English Abstract, TOC | `header1.xml` / `header2.xml` | `footer2.xml` / `footer3.xml` (`PAGE \* ROMAN`) | `fmt="upperRoman" start="1"` |
| **Section 3** | Main Chapters 1..N, References, CV, Acknowledgements, Declaration | Same `header1.xml` / `header2.xml` | `footer4.xml` / `footer5.xml` (`PAGE \* arabic`) | `fmt="decimal" start="1"` |

## 3. Clean Semantic Style IDs (`word/styles.xml`)

Note: `audit_thesis_docx.py` natively supports BOTH these clean semantic `styleId`s AND standard Word `Heading 1..3`/`Normal` OR legacy numeric template IDs (`01`..`13ref`).

| `styleId` | Role | Script Fonts (CN / Latin) | Size (`w:sz` / pt) | Paragraph Rules |
| :--- | :--- | :--- | :--- | :--- |
| `CoverTitle` | Cover Chinese title | 黑体 / TNR, **Bold** | `52` (`26 pt`) | `line="300 auto"`, `jc="center"` |
| `CoverMeta` | Cover field values | 仿宋 / TNR | `32` (`16 pt`) | `line="300 auto"`, `jc="center"` |
| `Heading1Chapter` | Chapter & top unnumbered headings | 黑体 / TNR, **Bold** | `32` (`16 pt`) | `pageBreakBefore`, `keepNext`, `keepLines`, `before=480 after=360`, `jc="center"`, `outlineLvl="0"` (`Ch.1` clears `pageBreakBefore`) |
| `Heading2Sec` | Section heading (`1.1`) | 黑体 / TNR, **Bold** | `28` (`14 pt`) | `keepNext`, `keepLines`, `before=480 after=120`, exact `20 pt` (`400`), `jc="left"`, `outlineLvl="1"` |
| `Heading3Subsec` | Subsection (`1.1.1`) | 黑体 / TNR, **Bold** | `26` (`13 pt`) | `keepNext`, `keepLines`, `before=240 after=120`, exact `20 pt`, `jc="left"`, `outlineLvl="2"` |
| `Heading4Para` | Level 4 heading (discouraged) | 黑体 / TNR, **Bold** | `24` (`12 pt`) | `keepNext`, `keepLines`, `before=240 after=120`, exact `20 pt`, `jc="left"` |
| `BodyTextZH` | Main body prose | 宋体 / TNR | `24` (`12 pt`) | `before=0 after=0`, exact `20 pt` (`400`), `firstLineChars="200"` / `firstLine="480"`, `jc="both"` |
| `AbstractBodyZH` | Chinese abstract prose | 宋体 / TNR | `24` (`12 pt`) | Exact `20 pt`, `firstLine="480"`, `jc="both"` |
| `AbstractBodyEN` | English abstract prose | TNR (`kern=0`) | `24` (`12 pt`) | Exact `20 pt`, `firstLine="240"`, `jc="both"` (`KEY WORDS:` sets `firstLine=0`) |
| `FigureCaption` | Caption below figure | 宋体 / TNR (**Unbolded**) | `22` (`11 pt`) | `keepLines`, `before=120 after=240`, `jc="both"`; bold only **图 N.M 短题名** at run level |
| `TableTitle` | Table title above table | 宋体 / TNR | `22` (`11 pt`) | `keepNext`, `before=240 after=120`, `jc="center"` (run-level bold title) |
| `TableCell` | Table cell text | 宋体 / TNR | `22` (`11 pt`) | `before=60 after=60`, `jc="center"` |
| `TableNote` | Notes/source below table | 宋体 / TNR | `21` (`10.5 pt`) | `before=120 after=240`, `jc="both"` |
| `BibEntry` | Bibliography entries | 宋体 / TNR | `21` (`10.5 pt`) | `before=60 after=0`, exact `16 pt` (`320`), `left="360" hanging="360"`, `jc="both"` |
| `TOCTitle` | `目  录` title | 黑体 / TNR, **Bold** | `32` (`16 pt`) | `pageBreakBefore`, `keepNext`, `keepLines`, `jc="center"` |
| `TOC1` / `TOC2` | TOC levels 1 & 2 | 宋体 / TNR (`TOC1` **Bold**) | `24` (`12 pt`) | Right dot leader tab `pos="8958"`, `TOC2 left="420"`, `line="360 auto"` |
| `HeaderBar` / `FooterPage` | Header & footer text | 宋体 / TNR | `18` (`9 pt`) | Center tab `4153` / right `8306`, `HeaderBar` bottom border `sz="6"` |
