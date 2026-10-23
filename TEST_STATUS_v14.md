# formatIQ TEST STATUS v14 (chat 14; update on top of TEST_STATUS_v12 and v13; read with HANDOFF_v26)

Method: sandbox (Linux, 4 GB RAM, 1 CPU). Node 22.22.2 and Node 18.20.8. Vitest 2.1.9, TypeScript 5.9.x, Vite 5.4.x, sax 1.x, zip.js 2.23.0.

## TESTED (really run, chat 14)
- watermark.ts (18808 bytes) + watermark.test.ts (14311 bytes): 35 of 35 pass on Node 22 AND Node 18, all on the first run; typecheck clean (strict, noUncheckedIndexedAccess, noUnused*); build ok (bundle 32 kB with xml_rewrite, package_merge and sax). No "any" type, no non-null "!", no non-ASCII characters (searched; the only search hit was the word "any" in a test title).
- Coverage: finding (text watermark with all fields, image watermark with relationship id, plain header and a shape with another id give nothing, several in order, the two Word markers); removing (whole run removed and other runs, text and paragraph properties kept, text and image together, header without watermark unchanged byte for byte, non-watermark picture kept); building text (opacity format like Word, diagonal text added after w:pPr with colour, opacity, rotation, id and re-read, horizontal has no rotation and a flatter box, size by text length within limits, escaping of text and font name and re-read of the original text, different ids through seed, eleven kinds of bad values give INVALID_OPTIONS); where the run goes (paragraph without properties, empty paragraph, empty pPr, only the first paragraph, header that starts with a table, header without a paragraph, missing namespaces added once with no duplicates, inserter reports); modes (inherit and keepSource unchanged, remove, text replaces an old one so exactly one remains, running twice is stable, result is well formed XML, damaged XML gives ZIP_READ_FAILED); image watermark (shape points at the relationship id, washout switch, replaces an old text watermark, bad picture settings); addImageRelationship (creates the file, keeps existing, first free id).
- Whole sandbox suite of chat 14: 213 of 213 pass in 7 files on Node 22 and Node 18 (zipio 28, package_merge 36, restyle 21, xml_rewrite 32, styles 22, page_setup 39, watermark 35). Files of the other chat (numbering_merge, validate, apply, apply_sections) were not loaded in this sandbox; their results are in TEST_STATUS_v12 and v13.
- Earlier results are unchanged (TEST_STATUS_v12, v13): page_setup 39, numbering_merge 29, validate 21, apply 21, apply_sections 11, and chats 7 to 10.

## CAVEATS (open)
1. WATERMARK DISPLAY IS UNVERIFIED IN WORD. The shape types, formulas, washout values, z-index and the size estimate were written from memory of what Word writes. The XML is well formed and re-reads in tests, but nobody has opened it in Word or LibreOffice. Try LibreOffice rendering to PDF if it can be installed; otherwise keep saying "not seen in Word".
2. watermark.ts is NOT wired into apply.ts yet (HANDOFF_v26 says how). page_setup.ts is not wired either (HANDOFF_v25).
3. Only the first paragraph of a header gets the run; watermarks stored as DrawingML are not found.
4. All earlier caveats from TEST_STATUS_v12 and v13 still stand (restyle test against the real paragraphs.ts, hand-made XML only, real Word files not tried).

## NOT TESTED
1. Real browser (zip.js Web Workers, sax in a browser bundle, file picker, showSaveFilePicker, dev server page): headless browser download blocked.
2. A true 1 GB docx (600 MB tested in chat 6); addBlob parts over 20 MB; zip.js passThrough.
3. Files saved by real Word, LibreOffice, Google Docs.
4. Cancel mid-write on a real disk stream; UTF-16 XML parts; 200 MB document.xml.
5. Lists set only by a style, localized style names, font-based code detection.
6. Wiring of page setup and watermark into apply, header/footer choices (5a), new section insertion, captions, cross-references, TOC, tables, images, notes, report, UI, template store, image reduction, batch (not built).
