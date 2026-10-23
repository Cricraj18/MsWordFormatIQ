# formatIQ TEST STATUS v13 (chat 13; update on top of TEST_STATUS_v12, which stays in place; read with HANDOFF_v21, v23, v24, v25)

Method: sandbox (Linux, 4 GB RAM, 1 CPU). Node 22.22.2 and Node 18.20.8. Vitest 2.1.9, TypeScript 5.9.3, Vite 5.4.21, sax, @zip.js/zip.js 2.23.0. The chat 13 sandbox was rebuilt from scratch: comment-free copies of types, xml, zipio, xml_rewrite, restyle, package_merge, numbering_merge, paragraphs, structure_rules, structure_detect, styles, template_extract (74880 bytes together, typecheck clean). page_setup.ts and validate.ts were NOT in this sandbox (the other chat tested them: see TEST_STATUS_v12).

## TESTED (really run, chat 13)
- src/core/apply.ts (21510 bytes) + apply_sections.ts (10658 bytes) with apply.test.ts (21 tests) and apply_sections.test.ts (11 tests): 32 of 32 pass on Node 22.22.2 AND Node 18.20.8; typecheck clean (strict, noUncheckedIndexedAccess, noUnused*); npm run build ok with a temporary entry (bundle 108 kB, 49 kB gzip; entry and index.html removed afterwards); no "any", no non-null "!", ASCII only (searched). Drive copies have the same byte sizes as the sandbox copies. The real output package was also dumped and read by eye (entries, document.xml, rels, content types, numbering, report): correct.
- Coverage apply.test.ts: package structural checks (every rels target exists, every r:id/r:embed resolves, every part has a content type, no override for a missing part); styles from detected roles (Heading1, Normal, ListParagraph) and direct formatting cleared (jc, spacing, rFonts, sz) or kept with directFormatting "keep"; paragraph text unchanged and same roles after detecting again; styles.xml, header, footer and settings from the template; template body image not copied; source header/footer references replaced in body sectPr and paragraph sectPr (ids tplrId10/tplrId11, titlePg from template); source landscape pgSz kept; source image clashing with a template image renamed (src_logo.png) and rel retargeted; hyperlink External kept; footnotes part and its own image (fn.png) followed; numbering merged (2 nums) and numId remapped in document and footnotes; numbering part added when the template has none (rel, override, numId unchanged); updateFields only with fields, in schema order; manual tag from a given detection (Title); output identical twice, source blob unchanged; missingStyles; report warning "not carried over"; errors INVALID_OPTIONS (scope selected), NOT_DOCX, TEMPLATE_INVALID, ZIP_READ_FAILED (damaged XML), CANCELLED (aborted signal).
- Coverage apply_sections.test.ts: model reader (first sectPr with references, empty model, titlePg val 0, sectPrChange ignored); fixer (replace references, titlePg schema order, xmlns:r added, source titlePg dropped, unmapped template id skipped, references inside sectPrChange removed without new ones, titlePg at end); addUpdateFields (before compat, at end, existing on, existing off turned on).
- One test mistake was fixed on the way (a self-closing element matches a name once, expected 1 not 2); no code bug found.
- Not redone: everything in TEST_STATUS_v12 (page_setup 39 tests, numbering_merge 29, validate 21, and all earlier chats). The Drive validate.ts in use is the 11565-byte one (TEST_STATUS_v12: 21 tests). A "validate 25 tests" figure that appeared in a duplicate TEST_STATUS_v12 written by this chat belonged to a 17162-byte validate.ts that is no longer in Drive src/core; ignore it.

## CAVEATS (open; TEST_STATUS_v12 caveats still apply)
1. The whole Drive suite has not been run together; restyle.test.ts etc. need one run against the real files in one sandbox.
2. apply was tested on hand-made zips only, not on real Word files.
3. apply does not call validate.ts; no test yet runs validateDocx on an apply output.
4. Scope "selected" (3c) is not built: apply throws INVALID_OPTIONS.
5. page_setup is not yet used by apply (source page size and margins are kept as they are).
6. All source sections get the same template header/footer references (per-section choices: 5a).
7. Output document.xml is gathered into Blobs in memory; a huge document.xml was not tested with apply.
8. Source docProps title is not copied; template docProps are used.

## NOT TESTED
As in TEST_STATUS_v12, plus: apply on real Word/LibreOffice/Google Docs files and its rendering in Word; apply with a huge document.xml; partial apply, header/footer per-section choices, watermark, captions, cross-references, TOC, tables, images, notes handling, report, UI, template store, image reduction, batch (not built).
