# Typography, Three-Line Tables, Biological Nomenclature & Citations

The numeric values and scientific examples below describe the bundled examples. Use the applicable school and discipline guidelines with an external template.

## 1. Script & Font Rules (`Err-8`)

- Chinese body prose: `宋体` (`SimSun`, `12 pt`), justified (`both`), exact `20 pt` line spacing, 2-character first-line indent (`480 dxa`).
- Headings: `黑体` (`SimHei`), bold, single ASCII space between numbering and title (`1.1 研究背景`, `Err-9`).
- All Latin letters, Arabic numerals, Greek symbols (`α`, `β`, `γ`, `μ`), `°C`, dashes, and English/formula parentheses MUST use `Times New Roman` (not SimSun/SimHei).
- Inside embedded vector artwork (`SVG`/`EMF`): sans-serif fonts (`Arial` / `黑体`, `6–9 pt` readable text, `10 pt` bold uppercase panel letters `A, B, C`).

## 2. Punctuation, Units, Dashes & Polarity (`Err-2-11-12`, `Err-4-5-6`, `Err-7`)

- **Chinese vs English punctuation**: Full-width punctuation in Chinese sentences (`，。；：（）“”、`); ASCII half-width in English prose, bibliography lists, and chemical formulas (`CO₂`, `(NH₄)₂SO₄`).
- **Enumerations vs Pairs**: Parallel genes/proteins use dunhao `、` (`TP53、MYC 和 KRAS`); equipment + vendor pairs use Chinese comma `，` (`设备型号，供应商`); multiple subpanels cited in Chinese prose use Chinese comma `，` (`图 1.1 A，B，F 和 H`), never dunhao (`图3.7A、B`).
- **Units**: Standardize `L`, `mL`, `µL`/`μL` (never `uL`, `ug`, `uM`), distinguish `mM` vs `mm`; single half-width space between number and unit (`50 mL`, `10 µL`), **no space** before `%` or `°C` (`95%`, `37°C`); roman lowercase `g` for grams, italic lowercase *g* for relative centrifugal force (`12,000 × g`).
- **Chapter 2 Time Units**: Choose **either** all Chinese (`周`, `天`, `小时`, `分钟`, `秒`) **or** all English abbreviations (`hr`, `min`, `sec`), never mixed.
- **Dashes**: Hyphen `-` for compounds/position numbers (`3-甲基吡啶`, `氧-氢键`); en dash `–` for numeric/page ranges (`285–298`); em dash `—` for breaks/structural bonds (`—COOH`); tilde `~` for high-energy bonds.
- **Sequence Polarity**: Label `5′-` / `3′-` on nucleic acids and `N-` / `C-` on peptides/proteins.
- **Acronyms (`Layout-7`)**: Expand once in Chinese Abstract, once in English Abstract, and once at first mention in Main Body; never repeat full expansions within the same scope.

## 3. Three-Line Tables & Figure Captions (`Tbl-3Line`, `Fig-1..10`, `Acad-2`)

- **Three-Line Table Borders**: Top = `1.5 pt` (`w:sz="12"`), Bottom = `1.5 pt` (`w:sz="12"`), Header row bottom = `0.75 pt` (`w:sz="6"`), `left` / `right` / `insideH` / `insideV` = `nil`. Header row uses `<w:tblHeader/>` and `<w:cantSplit/>`.
- **Table Title (`TableTitle`, `11 pt` centered above table)**: Bold only **表 N.M 短题名**; notes (`TableNote`, `10.5 pt`) below table.
- **Figure Caption (`FigureCaption`, `11 pt` unbolded style default, justified below figure)**: Apply run-level bold ONLY to **图 N.M 短题名**, followed by regular detailed explanations of uppercase subpanels (`A`, `B`, `C`), sample size `n`, statistical test + *P*-value, and scale bars.
- **Adapted Literature Figures**: Place citation superscript/tag directly after caption title, and append `修改自 Author et al., Year。` at the **very end** of the caption.

## 4. Citation Systems (`--citation-style`)

### Option A: `nature` (Default)
- In-text: superscript numbers **before** punctuation (`……已有报道¹,²。`), zero exposed `ADDIN EN.CITE` fields.
- Reference list (`BibEntry`, `10.5 pt`, `16 pt` exact line spacing, justified, hanging indent):
  1. Space between initials (`S. F.`);
  2. `&` before final author (never `and`);
  3. Sentence-case article title;
  4. *Italic* title-case abbreviated journal name (`Nat. Rev. Mol. Cell Biol.`, exceptions: `eLife`, `iScience`);
  5. **Bold** upright volume number;
  6. En dash `–` in page range;
  7. `(Year)` with **no trailing period** and no DOI unless ahead-of-print;
  8. English references sorted **before** Chinese references.

### Option B: `gb7714`
- In-text: sequential `[1]`, `[1,2]`, `[1–3]` markers.
- Reference list (`BibEntry`): numbered `[1]`, uppercase or standard author list (`et al.` / `等`), article title followed by mandatory document type tag (`[J]`, `[M]`, `[D]`, `[C]`, `[R]`, `[EB/OL]`), journal name upright, year, volume(issue): start–end page.
