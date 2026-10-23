# formatIQ HANDOFF v17 (COMPLETE FEATURE REGISTER; short file on top of v16)

READ ORDER: (1) HANDOFF_v6.md (2) HANDOFF_v8.md (3) HANDOFF_v9.md (4) HANDOFF_v15.md (5) HANDOFF_v16.md (6) This file (7) TEST_STATUS_v6.md (8) REVIEW_v1.md and REVIEW_v1_addendum.md.
Trust order if files disagree: v17, v16, v15, v9, v8, TEST_STATUS_v6, REVIEW, v6.
CLEAN-UP RULE change: keep HANDOFF_v6, v8, v9, v15 AND v16 in place. Everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong" as before. progress_v11 and TEST_STATUS_v6 are still the highest (no code was written in this step).
Status unchanged: step 3a done. NEXT is 3b apply.ts. The user said: include ALL features, nothing left out. This file is the full register. Build in the tier order below; a lower tier never blocks a higher one. When the chat has room at its end, consolidate v6, v8, v9, v15, v16, v17 into ONE self-contained HANDOFF (then move the six old ones to "wrong"): reading six files every chat wastes the context.

## KEY DESIGN RISK FOUND (must be solved in 3b, not later)
Changing w:pStyle alone does NOT change how a paragraph looks if the paragraph has DIRECT formatting (fonts, sizes, spacing, indents, alignment set on w:pPr or w:rPr). Also numbering ids (w:numId) in the source point to the SOURCE numbering.xml, which the template package does not have; relationship ids (images, hyperlinks, footnotes) differ between packages; media and content types must be copied. So 3b needs: (a) a direct-formatting policy option: "clear" (default: remove paragraph-level spacing/indent/alignment/font/size/color and run-level font/size/color, KEEP bold, italic, underline, strike, sub/superscript, highlight, hyperlinks, fields, tabs, breaks, images), or "keep"; (b) numbering policy: map source lists to template list styles/numbering definitions, or keep the source numbering.xml definitions merged into the template numbering.xml with new ids; (c) relationship and content-type merge for everything poured in from the source; (d) a validator (validate.ts) that re-reads the output and checks every r:id resolves, every part has a content type, bookmark ids are unique, w:sectPr are well formed, styles referenced exist in styles.xml. Test with files made by the docx npm package AND hand-made edge cases.

## FEATURE REGISTER (every item must end up built, tested in sandbox, and listed in TEST_STATUS)
Tier 1, core (nothing else works without these):
- 3b apply.ts: template package as base; paragraph styles from styles.ts; direct-formatting policy; numbering merge; rels/content types/media merge; unknown parts preserved (comments, footnotes, endnotes, SmartArt, equations, custom XML, tracked changes). Paragraph index MUST match readParagraphs; test by formatting then detecting again.
- 3c partial apply (scope "selected", FormatPart list: styles, headersFooters, toc, crossReferences, sections).
- validate.ts output integrity check (above).
- 5a headers/footers inheritance (v16 item 1), page numbers continuous across sections.
- 5b page_setup.ts + sections: A4, Letter, A3 (+ Legal, A5, Tabloid), custom size, portrait/landscape, margins incl. header/footer distance and gutter, mirror margins, per-section or all, section breaks kept, new section insertion (landscape for wide tables, tag-only landscape with Next-Page breaks and swapped page size per the v6 spec).
Tier 2, document features:
- 5c watermark: inherit, remove, keep source, add text, add image (v16 item 5).
- 4a captions: SEQ Figure and SEQ Table fields with bookmarks, caption style, "keep with next" so a caption stays with its figure or table. 4b cross-references: REF fields to those bookmarks; STYLEREF where useful.
- toc.ts: Table of Contents, List of Figures, List of Tables as native fields with the template's TOC styles; update-fields-on-open (settings.xml w:updateFields) plus a visible warning that fields may look empty until updated; rebuild an existing TOC instead of duplicating it.
- Tables: apply the template's table style, repeat header row, fit width to the text width, cell margins and alignment, no row splitting across pages option, table captions position (above), wide-table landscape option.
- Images: fit to the text width without upscaling, centre, keep with caption, alt text preserved (add a warning list of images without alt text), anchored vs inline kept as is.
- Lists: bullets and numbering mapped to template list styles; multilevel numbered headings (1, 1.1, 1.1.1) linked to heading styles; consistent indents.
- Footnotes and endnotes: apply the template's Footnote Text/Reference styles; keep numbering.
- Page numbering: format (arabic, roman), restart per section, "Page X of Y" with PAGE and NUMPAGES fields; cover page without a number (titlePg).
- Cover/title page: built from the template's Title/Subtitle styles when a title role is detected; option on/off.
- Fonts and theme: come with the template package (theme1.xml, fontTable, styles); clear conflicting direct fonts per the policy; embedded fonts of the template kept.
- Hyperlinks and character styles: Hyperlink style from the template; bookmarks, internal links and fields kept.
- Track changes and comments: options keep (default), accept all, reject all, remove comments; never lose them silently.
- Headers/footers extras: different first page, even/odd, section-specific, linked-to-previous; fields inside kept.
- Document properties: core.xml title/author kept or replaced; language and proofing (w:lang) kept; right-to-left (w:bidi) paragraphs and complex-script text kept intact; compatibility mode of the template kept.
- Equations (OMML), text boxes, SmartArt, charts, embedded objects, content controls, form fields, sections inside tables: pass through untouched and never counted as paragraphs the wrong way.
Tier 3, product shell:
- 6 manual_mark.ts: manual tagging of low-confidence paragraphs (bulk actions, jump to next unsure, filter by role).
- 7 web UI (phone friendly, simple language): pick source and template, options (v16 list), preview of detected structure, progress and Cancel, error messages in plain words, drag and drop, remember last options, privacy note ("files never leave your computer"), keyboard and screen reader basics, Web Worker with streaming output (showSaveFilePicker with Blob fallback).
- 8 template_store.ts: IndexedDB store, import/export template files, template summary card (styles found, page size, margins, header/footer yes/no, watermark yes/no, TOC yes/no, missing styles warning from missingStyleNames).
- report.ts: before/after change report (what was restyled, what was skipped, warnings, low-confidence list) as an on-screen list and downloadable text/HTML; the original file is never modified (output is a new file), so "undo" = the original.
- 9 image_reduce.ts (optional module, OFF by default): Original / Light (~2400 px, Q90, ~1 MB) / Balanced (~1600 px, Q90, ~500 KB); pixels only, never upscale, keep displayed size and crop, min 150 DPI, skip line art and transparent images, one image at a time in a worker, progress and Cancel.
- batch.ts: format several files with one template in a queue, one output each, one combined report.
- 10 deploy guide (static hosting, no server) and a short user guide.
Tier 4, decide with the user before building (flag clearly):
- Legacy .doc (binary), .rtf, .odt as sources: needs a converter that is not realistic in a browser-only tool; recommended answer: the user saves as .docx first. .docm is already allowed by the spec.
- Google Docs as a source: recommended answer: download as .docx first.
- PDF input, OCR, AutoCAD: explicit NON-GOALS in the v6 spec. Keep out unless the user changes the spec in writing.

## DIRECTION CHECK GIVEN TO THE USER IN CHAT 10
Foundation matches the spec and is solid (zip, safe XML, template reading, detection, XML rewrite, style mapping, 160+ tests). Biggest risks: (1) nothing yet seen on real Word/LibreOffice/Google Docs files or in a real browser; (2) the template-package merge (3b) is the hardest part; (3) scope is large, so tiers matter; (4) handoff files are heavy, so consolidate. Suggested but optional help from the user (NOT testing): put 3 to 5 real .docx files (a messy report, a thesis with figures and tables, a file with tracked changes and comments, a landscape/section-heavy file) and 1 to 2 real templates in a Drive folder named "samples" inside FORMAT_IQ. If a "samples" folder exists, read it and use those files in tests (copy through download_file_content, base64).

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER. Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested. After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
