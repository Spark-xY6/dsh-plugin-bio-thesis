# 65-Item Self-Check Matrix & Automated Audit Mapping

Mapping between the **65 checklist items** in `resources/bio-thesis-self-check-65.docx` (`20 + 9 + 10 + 9 + 17`) and the checks in `scripts/audit_thesis_docx.py`. Select entries according to the school instructions and the work being submitted. Mark irrelevant items as not applicable. Laboratory nomenclature, quantitative statistics, citation conventions and preset page dimensions apply only where the discipline or template requires them. The defaults below describe the bundled example. With `--template`, document geometry and section settings follow the supplied template. Its style definitions are compared in `Layout-TemplateStyles`. School-specific checks marked `INFO` require the applicable institutional rule or manual verification.

Items marked **Manual / Visual QA** require human or live Word/DSH `office_render` inspection and are never claimed as verified by XML parsing alone.

## I. 学术规范与科学表述 (Items 1–20)

| # | Requirement Summary | Automated Check ID / QA Mode |
| :---: | :--- | :--- |
| 1 | Identify sources and avoid unattributed copying | Manual QA |
| 2 | Adapted literature figures unified in style, cited after title, `修改自 Author et al., Year。` placed at **caption end** | Auto `Acad-2` + Visual QA |
| 3 | Credit external figures, tables and source materials | Manual QA |
| 4 | Formal Chinese scientific grammar, non-colloquial prose | Manual QA |
| 5 | English abstract & prose free of grammar/spelling defects | Manual / Word Proofing |
| 6 | Consistent plotting palette, font scale, and stroke style | Visual QA |
| 7 | Consistent narrative tone and terminology across chapters | Manual QA |
| 8 | Claims and external methods properly cited | Manual QA |
| 9 | Priority claims supported by literature & evidence | Manual QA |
| 10 | Collaborative credits clear (`final` names vs `blind` anonymous lab phrasing) | Auto `Acad-10-Blind` + Manual QA |
| 11 | Zero typos; sequence polarity marked (`5′-`/`3′-`/`N-`/`C-`) | Manual QA |
| 12 | When using statistical tests, document the method, sample size and applicable statistics | Manual QA |
| 13 | If required by the chosen writing convention, avoid first-person pronouns (`我`/`我们`); single consistent self-term (`本论文`/`本研究`/`本工作`) | Auto `Acad-14` |
| 14 | Genes italic, proteins roman, mutants lowercase italic | Auto `Acad-15-16` (`--genes`) |
| 15 | Species binomials italic; restriction/DNA enzymes italic w/ space before Roman numeral (`EcoR I`, `Taq I`) | Auto `Acad-15-16` (`--species`) |
| 16 | Chinese enumerations use `、`; English enumerations use `,` | Auto `Err-2-11-12` + Manual QA |
| 17 | Reference & chemical formula parentheses use ASCII `()` | Auto `Err-2-11-12`, `Ref-2-7` |
| 18 | Correlation distinguished from causation | Manual QA |
| 19 | Novel genes/molecules follow field nomenclature | Manual QA |
| 20 | Section `1.2` presents prospective aims & significance, not a retrospective summary | Manual QA |

## II. 页面、分节与整体排版 (Items 1–9)

| # | Requirement Summary | Automated Check ID / QA Mode |
| :---: | :--- | :--- |
| 1 | A4 & margins match preset (`3.0/2.5/2.6/2.6 cm`, `mirrorMargins`, `evenAndOddHeaders`) | Auto `Layout-Settings`, `Layout-1`, `Layout-SecNum` |
| 2 | Chapter & front/back matter order matches degree/milestone mode | Bundled starter scaffold or supplied template + manual review of the finished document |
| 3 | Script fonts unified (`宋体`/`黑体`/`Times New Roman`) | Auto `Err-8` |
| 4 | Font sizes conform to `styles.xml` hierarchy | Defined in master styles |
| 5 | Exact line spacing (`20 pt` body / `16 pt` refs) & indents unified | `Layout-BodySpacing` for recognized body styles, `Ref-9` for bibliography line spacing/alignment, manual verification for other styles and indents |
| 6 | English letters, numerals, Greek (`α`, `β`, `μ`), and `°C` use Latin font (`TNR`/`Arial`) | Auto `Err-8` |
| 7 | Acronyms expanded once in Chinese Abstract, English Abstract, and Main Body | Auto `Layout-7` (`WARN`) |
| 8 | No widow/orphan headers or broken split tables | Visual / `office_render` QA |
| 9 | Abstracts, body prose, figure captions, and bibliography justified (`both`) | Auto `Layout-9`, `Ref-9` |

## III. 插图与三线表规范 (Items 1–10)

| # | Requirement Summary | Automated Check ID / QA Mode |
| :---: | :--- | :--- |
| 1 | Sequential chapter/appendix numbering (`图 1.1`, `表 2.1`, `A.1`) | Auto `Fig-1` |
| 2 | Figures/tables and captions/notes remain together cleanly | Visual / `keepNext` + `cantSplit` |
| 3 | Every figure and table cited in main body prose | Auto `Fig-3-8` |
| 4 | Caption (`FigureCaption`) & table title (`TableTitle`) selective bolding; 3-line borders (`1.5 pt / 0.75 pt / nil`) | Auto `Fig-4`, `Tbl-3Line` |
| 5 | In-figure text `6–9 pt` readable; unified scale bars | Visual QA |
| 6 | Balanced figure density; no oversized/overflowing panels | Visual QA |
| 7 | High-resolution / crisp vector output (not accidentally rasterized) | Auto `Pkg-Media` + Visual QA |
| 8 | Every figure/table substantively interpreted in text | Auto `Fig-3-8` + Manual QA |
| 9 | Panel labels uppercase (`A, B, C`; recommended `10 pt Arial`) | Auto `Fig-9-10` + Visual QA |
| 10 | Detailed explanatory captions (no single-line title captions) | Auto `Fig-9-10` |

## IV. 文献引用与著录规范 (Items 1–9)

| # | Requirement Summary | Automated Check ID / QA Mode |
| :---: | :--- | :--- |
| 1 | Citation placement & format match active `--citation-style` (`nature` superscript before punctuation vs `gb7714` `[1]`); no raw `ADDIN EN.CITE` | Auto `Ref-1`, `Ref-2-7` |
| 2 | Nature 7-point check: spaced initials, `&`, sentence-case title, *italic* journal, **bold** vol, en-dash `–`, `(Year)` w/o trailing `.` | Auto `Ref-2-7` (`nature`) |
| 3 | Redundant DOIs removed from formally published entries | Auto `Ref-2-7` |
| 4 | Nature titles sentence-case | Auto `Ref-2-7` |
| 5 | GB/T 7714 document type markers (`[J]`, `[M]`, `[D]`, etc.) & metadata present | `Ref-2-7` screens type markers and missing journal/year patterns. Complete metadata and publisher rules require manual review. |
| 6 | Journal casing & English-before-Chinese sorting consistent (`eLife`/`iScience`) | Auto `Ref-2-7` |
| 7 | No duplicate bibliography entries | Auto `Ref-2-7` |
| 8 | Bidirectional closure between body citations and bibliography | Auto `Ref-8` |
| 9 | `BibEntry` exact `16 pt` line spacing, justified, hanging indent | Auto `Ref-9` |

## V. 专业体例与常见排版错误 (Items 1–17)

| # | Requirement Summary | Automated Check ID / QA Mode |
| :---: | :--- | :--- |
| 1 | TOC fields refreshed (`F9`) with unbroken dot leaders & bold `TOC1` | Auto `Layout-TOC` + Word `F9` |
| 2 | No half-width/full-width punctuation mixing in body | Auto `Err-2-11-12` |
| 3 | Re-drawn/cropped literature graphics clean of stray artifacts & cited | Visual QA + `Acad-2` |
| 4 | Chemical formulas/ions formatted cleanly (`CO₂`, `H₂O`, `pH`) | Auto `Err-4-5-6` |
| 5 | Units standardized (`L`, `mL`, `µL`, never `uL`/`ug`/`uM`; `mM` vs `mm`) | Auto `Err-4-5-6` |
| 6 | Single ASCII space between number & unit (`%` and `°C` excepted) | Auto `Err-4-5-6` |
| 7 | Chapter 2 time units uniformly Chinese OR English, never mixed | Auto `Err-7` |
| 8 | All Latin/numeric runs inside Chinese headings & body use English fonts | Auto `Err-8` |
| 9 | Single half-width space between heading number and title | Auto `Err-9` |
| 10 | Long headings break at semantic boundaries | Visual / `office_render` QA |
| 11 | Duplex page-numbering / header continuity across sections | Auto `Layout-SecNum` + Visual QA |
| 12 | Equipment/vendor pairs & subpanel enumerations use Chinese comma `，` (`图 1.1 A，B`) | Auto `Err-2-11-12` |
| 13 | Genes, species, and restriction enzymes properly italicized | Auto `Acad-15-16` |
| 14 | Bilingual inline spacing consistent without stray double spaces | Manual QA |
| 15 | Upright `g` (gram) vs italic *g* (centrifugal force) | Manual QA |
| 16 | Accurate `-` hyphen, `–` en dash, `—` em dash, and `~` high-energy bond | Manual + `Ref-2-7` |
| 17 | Maps containing borders strictly use officially approved standard maps | Manual QA |

## VI. Additional Package & Scope Checks

`Pkg-CRC`, `Pkg-XML`, `Pkg-Rels`, `Pkg-Clean`, `Pkg-Media` (baseline SHA-256 verification), and `WordCount` (`INFO`/`WARN` by mode target).
