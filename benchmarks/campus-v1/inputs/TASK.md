# Campus archive • synthetic Word task

Only the files in this input folder and the seven stated requirements govern this assignment. All names, text and graphics are synthetic. Use the files locally. There are two tasks. Complete both in the output directory specified by the evaluator. Keep all input files intact.

## 1. Audit all eight submitted drafts

Audit A01.docx through A08.docx against school.docx. The campus accepts ordinary wording changes. Report every violated requirement below using the seven exact labels. Include an empty list when a draft satisfies all seven. Apply each label independently. Titles and author placeholders in the school template are already filled in the submitted drafts.

| Label | Requirement |
| --- | --- |
| PAGE | Both sections use portrait Letter (12240 × 15840 twips). Each section's top, right, bottom and left margins must be 1740, 1320, 1500 and 1540 twips, respectively. |
| NUMBER | The first section starts with uppercase Roman I and the second restarts with decimal 1. The default footer of **each** section must contain a live PAGE field. A typed digit is insufficient. |
| STYLE | The paragraph style named “Campus body” (ID BodyCampus in the template) must define/inherit a size of 13 pt. |
| DIRECT | The paragraph “Anchor paragraph.” must use the campus body style with no local font-size override. A change to the style definition alone belongs solely under STYLE. |
| MEDIA | Embedded media must have the same package-relative filenames and bytes as school.docx. |
| IDENTITY | The forbidden, case-insensitive identity token is EXAMPLE-Q17. Search ordinary document text (including across adjacent runs), headers, footers, comments, file properties and XML attributes. Images need only the separate byte comparison above. |
| TOC | The document body must contain a live TOC field. Static directory text or a different field does not satisfy this requirement. |

Write exactly one UTF-8 JSON file named result.json, in this form: {"cases":{"A01":[],"A02":[],"A03":[],"A04":[],"A05":[],"A06":[],"A07":[],"A08":[]}}. Replace the example arrays with the exact labels you find; arrays can contain multiple labels. A note in prose does not substitute for the JSON file.

## 2. Generate a blind-review draft

Use school.dotx as the source. school.docx is the equivalent macro-free document version, provided for tools which require DOCX. Produce generated.docx as a standard editable Word document. Fill the split English-title placeholder with **Forms of Evidence**. The supplied author is **EXAMPLE-Q17**. Replace the author marker with a visible anonymous designation (e.g. Anonymous, 匿名, or 【匿名评审隐去】); keep the author label and omit the identity token throughout XML and document properties.

Keep the original ordinary text (including “1  Materials”, “Anchor paragraph.” and “Unchanged ordinary paragraph.”), both sections and their original paper, orientation, margins and numbering, both live PAGE footers, the live TOC, the 13-pt campus body style applied to the anchor, the embedded image and its image instance, and the header “Campus archive”. The evaluator checks these 12 aspects independently against the template and requirements. Filenames, bytes and routine XML serialisation are free to change where the requirements only ask for preserved meaning.
