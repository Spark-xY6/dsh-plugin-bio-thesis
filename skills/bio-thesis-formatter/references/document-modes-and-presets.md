# Document Modes, Institutional Presets & Blind Review Switching

Use the current school template with `init --template` and `audit --template` when one is available. JSON presets configure the bundled starter and optional school-specific checks.

## 1. Built-in Presets & Institutional Overrides

| Preset (`--preset`) | Citation Style | Target Use Case |
| :--- | :--- | :--- |
| `cn-lifesci-default` | `nature` | Chinese life sciences, biology, and biomedical dissertations/reports |
| `cn-gb7714-general` | `gb7714` | Chinese-language theses in disciplines that require a sequential numeric reference style (`[1]` + `[J]`/`[M]`/`[D]`) |

### Customizing for Any University (`--profile-config <file.json>`)

Start from `presets/custom-university.example.json` in the repository root and override only the fields your graduate school modifies:

- `university_name` / `school_name` (or pass `--university` and `--school` on the CLI)
- `citation_style`: `"nature"` or `"gb7714"`
- `page.width_dxa`, `page.height_dxa`, `page.margins_cm` (`top`, `bottom`, `left`, `right`, `header`, `footer`, `gutter`) — `init_thesis_docx.py` converts updated `margins_cm` automatically across all 3 sections, and `audit_thesis_docx.py` audits against those exact dimensions
- Optional custom emblem on cover: pass `--logo path/to/emblem.png` (or `.jpg`) to `init`

## 2. Five Document Modes (`--mode`)

| Mode | Cover & Header Focus | Copyright & Declaration | Chapter Skeleton & Word Target |
| :--- | :--- | :--- | :--- |
| `phd-thesis` | Doctoral dissertation | Included (`S1` + `S3` placeholders) | Ch1–Ch5 + References + CV + Acknowledgements + Declaration |
| `master-thesis` | Master thesis | Included (`S1` + `S3` placeholders) | Ch1–Ch5 + References + CV + Acknowledgements + Declaration |
| `proposal` | Graduate thesis proposal (开题报告) | Omitted | Ch1 背景综述 (`1.2 本研究的目的、主要内容和意义`), Ch2 研究材料与方法, Ch3 前期进展, Ch4 后续计划与方案设计 (`8,000–20,000` suggested) |
| `midterm` | Midterm / annual evaluation report | Omitted | Uses `proposal` chapter skeleton (`>= 5,000` suggested) |
| `rotation` | Lab rotation report | Omitted | Ch1 轮转背景 (`1.2 本轮转课题的目的与内容`), Ch2 材料与步骤, Ch3 结果与讨论, Ch4 收获与建议 (`>= 2,000` suggested) |

## 3. Double-Blind Anonymization (`--review-stage blind`)

| Target | `final` Stage | `blind` Stage |
| :--- | :--- | :--- |
| Cover table (`姓名`, `学号`, `导师姓名`) | Author-supplied values | Replaced with `【匿名评审隐去】` |
| English abstract attribution | `<Author> (<Major>)` / `Directed by <Advisor>` | `Anonymous Author (<Major>)` / `Directed by 【匿名评审隐去】` |
| CV bio & publications | Full author & publication lines | Identity hidden; the placeholder publication line becomes a note about anonymous attribution |
| Acknowledgements | Full acknowledgements section | Replaced with `【匿名评审版本隐去致谢，终稿恢复。】` |
| Sensitive scan (`Acad-10-Blind`) | Skipped | Pass `--sensitive-names "Name1,Name2"` during audit to screen recognized XML content and file properties. Inspect images, signatures and embedded files separately |

## 4. `1.2` Heading & Consistent Self-Term (`Acad-14`)

- Degree theses (`phd-thesis` / `master-thesis`): `本论文`
- Proposals / Midterm (`proposal` / `midterm`): `本研究`
- Lab rotation (`rotation`): `本工作`
- First-person pronouns (`我们`) outside neutral declaration templates fail `Acad-14`.
