# Report artifact contract

## Authority and retained files
- Primary retained source: C:\Users\Dogs\Desktop\GraduationDesign\02_Document\.work\original_202310232005.docx
- SHA-256: f6df12fcb29d9b0d968dd0311bdd591b2354a98c05067c82814e9c7db0d85e0e
- Source rendered read-only by Microsoft Word to .work/original.pdf: 10 pages, 2 sections.
- Secondary design reference: C:\Users\Dogs\Desktop\GraduationDesign\02_Document\开题报告-姓名（20261115）.docx; SHA-256 9152605267ad26626b069481f28424c0aa31452f99c4eb84a61d14355dd97e3a; 18 pages rendered to .work/template.pdf.
- PDF example: 201910121236_郎玮亮_开题报告.pdf, 9 pages, read for section sequence and density.
- Baseline package inventory: .work/package-before.json. Preserve every existing part byte-for-byte except word/document.xml and word/_rels/document.xml.rels. Add only two figure PNGs and necessary image/hyperlink relationships. Existing images, styles, numbering, customXml, settings, theme and metadata are preserve-only.

## Page and type system
- Both sections A4 11906 x 16838 twips, portrait. Margins top/bottom 1440 and left/right 1800 twips. Header distance 851, footer 992 twips. Single column, 425 twip spacing. Existing section break and cover untouched.
- Preserve cover university wordmark/logo, centered title, metadata table, name 蒋以怀, student number 202310232005, college 物流工程学院, major 自动化, class 自动化231, advisor 燕存良 and cover date 2026年11月15日.
- Body follows source: 宋体 12 pt, Times New Roman for Latin, 1.5 line spacing, first-line 480 twips. Heading 1 uses original style id 1, 12 pt bold, black, source-derived before/after spacing. Subheadings 12 pt bold, black. References 10.5 pt, hanging 360 twips, approximately 13 pt baseline. Captions 10.5 pt centered. Native OMML equations.
- No new page furniture. Tables keep original black single-line borders, cell grid, widths, cover/table layouts, existing teacher and review spaces. Narrative cells aligned top and allowed to split across pages; font is not reduced to fit.

## Slot map and content flow
Stable locations are /w:document/w:body/w:tbl[2]/w:tr indexed 1-based.
- tr[1]: preserve report title and project name.
- tr[2]: rewrite background/significance, remove filler blanks, retain prompt and heading pattern.
- tr[3]: rewrite research review, methods, two original diagrams, native equations, twenty references with DOI hyperlinks. Remove unrelated wind-power content.
- tr[4]: rewrite concrete tasks, measurable planned objectives, key technical issues. No fabricated results.
- tr[5]: rewrite thesis outline; preserve section purpose.
- tr[6:7]: preserve schedule title and headings.
- tr[8:16]: fill nine schedule task cells; retain date ranges except invalid 2027-02-29 corrected to 2027-02-28.
- tr[17:29]: preserve blank schedule rows as source layout.
- tr[30]: keep teacher-only type choices, remove inherited checked theory selection to leave choices open.
- tr[31:32]: preserve teacher source and opinion/signature fields without authoring an opinion.
- body/tbl[3] and review heading: preserve all review-only fields.
- All body paragraphs outside the specified slots, including title/footer/section paragraphs, remain unchanged.

## Figures and citations
Two figure PNGs at 5,050,000 EMU width (about 14.0 cm), inline, centered, with caption kept adjacent. New image relationships only. Academic bibliography contains exactly 15 Chinese and 5 English journal papers from 2022-2025. Each is cited in the research text and has its verified DOI. Source metadata is stored in .work/references.json.

## Fidelity and delivery gates
- Retained source and secondary DOCX hashes unchanged.
- Unchanged ZIP parts match original hashes; no loss of opaque parts or relationships.
- Page geometries and section count match. Cover, formal report borders and reviewer page remain recognizable.
- No wind-power template prose or copied advisor opinion in authored content. All stated metrics are targets with test conditions.
- Packaged render_docx.py attempted and failed because soffice.exe is unavailable. No bundled LibreOffice exists on this Windows runtime. Use read-only Microsoft Word ExportAsFixedFormat as rendering fallback, rasterize with bundled Poppler, inspect every final page PNG and iterate. Never save the source or final document through the renderer.
- Deliver only the requested DOCX; PDF/PNGs are internal QA artifacts.

## Final verified layout refinements
- Final report: 15 pages, 2 unchanged sections. Introductory cell distributed across source-derived continuation rows to prevent a blank first body page.
- Schedule starts as a separate source-derived table on page 13; two unused spacer rows removed. Thesis outline fits page 12. Advisor form page 14 and review title/form page 15.
- Review title kept with form, with original negative text-position override removed. Inherited teacher-only ticks cleared. No instructor opinion or review result authored.
- All 15 final PNGs in .work/final-render inspected. Original and reference DOCX files retained. Package checks passed.
