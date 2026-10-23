# formatIQ TEST STATUS v12 (chat 12; replaces TEST_STATUS_v11; read with HANDOFF_v21, v23, v24)

Method: sandbox (Linux, 4 GB RAM, 1 CPU). Node 22.22.2 and Node 18.20.8. Vitest 2.1.9, TypeScript 5.9.x, Vite 5.4.x, sax 1.x, zip.js 2.23.0, docx npm package (fixtures).

## TESTED (really run)
Chat 12 (this chat):
- page_setup.ts (16829 bytes) + page_setup.test.ts (16895 bytes): 39 of 39 pass on Node 22 AND Node 18; typecheck clean (strict, noUncheckedIndexedAccess, noUnused*); build ok (bundle 28 kB with xml_rewrite and sax). All 39 passed on the first run. No "any", no non-null "!", no non-ASCII characters (searched). Coverage: units and presets (mm and inches, preset table is portrait, default options); reading sections (paragraph and body sections in order, landscape from orient and from width > height, missing pgSz/pgMar, sectPrChange is not a section, no sections); portraitSize and checkMargins (template, A4 fallback, presets, custom, custom limits, margin problems); planSection (template default, preset, landscape swap, portrait forces portrait, source landscape stays landscape, landscape template, custom and empty margins, selection, INVALID_OPTIONS with section number); applying (size, margins and paper code removal, orient add and remove, other section children untouched, creation of missing pgSz/pgMar in the right place, empty sectPr, old pgMar kept when no margins, only selected sections changed and others copied exactly, source landscape kept, sectPrChange record untouched, own prefix, text and tables unchanged, stable when run twice, no sections copied unchanged, invalid custom size); building blocks; streaming (20000 sections in 64 KB pieces, about 0.8 s).
- Full sandbox suite of chat 12: 178 of 178 pass in 6 files on Node 22 and Node 18 (zipio 28, package_merge 36, restyle 21, xml_rewrite 32, styles 22, page_setup 39). numbering_merge and validate were not loaded in this sandbox.
Chat 11 (other chat, from TEST_STATUS_v11, not redone): numbering_merge.ts + test 29 of 29 and validate.ts + test 21 of 21 on Node 22 and Node 18, typecheck clean, build ok. numbering_merge coverage: split, ids, picture bullets, id shifts and abstractNumId update, template children byte for byte, source list-level pStyle links removed with a warning, no template numbering, high template ids, numIdMacAtCleanup, namespace and mc:Ignorable union, style-link warning, deterministic, unique ids, damaged XML and DOCTYPE, numId remap. validate coverage: scanner (ids, styles, lists, bookmarks, sections, 7-character pieces, nested sectPr in sectPrChange, damaged XML), every problem code, plus ONE real .docx made by the docx library: zero errors.
Chats 7 to 10 (not redone): xml 12, template_extract 27, paragraphs/structure 56, zipio 28, package_merge 36, restyle 21, xml_rewrite 32, styles 22.

## CAVEATS (open)
1. restyle.test.ts still needs a run against the real Drive paragraphs.ts and xml.ts (sandbox used a partial paragraphs.ts copy and a comment-free xml.ts). Do it when apply.ts is built.
2. numbering_merge was tested on hand-made numbering XML, validate.ts on one docx-library file; page_setup on hand-made XML. None on a Word-saved file.
3. Source list definitions using w:styleLink / w:numStyleLink only produce a warning.
4. When the template has no numbering part but the source has one, apply.ts must add the part, relationship and Override (HANDOFF_v23).
5. page_setup: margins do not rotate with the page; no function yet to insert a NEW section break.
6. validate.ts does not check relationship TYPES against content types and does not check field codes.

## NOT TESTED
1. Real browser (zip.js Web Workers, sax in a browser bundle, file picker, showSaveFilePicker, dev server page): headless browser download blocked.
2. A true 1 GB docx (600 MB tested in chat 6); addBlob parts over 20 MB; zip.js passThrough.
3. Files saved by real Word, LibreOffice, Google Docs.
4. Cancel mid-write on a real disk stream; UTF-16 XML parts; 200 MB document.xml.
5. Lists set only by a style, localized style names, font-based code detection.
6. apply.ts (not built yet); header/footer inheritance, watermark, new section insertion, captions, cross-references, TOC, tables, images, notes, report, UI, template store, image reduction, batch (not built).
