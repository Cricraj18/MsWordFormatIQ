# formatIQ HANDOFF v34 (self-contained; short update on top of HANDOFF_v21, v23 to v33)

READ ORDER: (1) HANDOFF_v21. (2) v23 to v33. (3) This file. (4) TEST_STATUS_v12 to v21. (REVIEW and progress files: optional history.)
Trust order if files disagree: v34, v33, v32, ... v23, v21, TEST_STATUS_v21 down to v12, REVIEW.
CLEAN-UP RULE: keep HANDOFF_v21 and v23 to v34 in place, keep TEST_STATUS_v12 to v21 in place, keep the highest progress file. Everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda). Never delete.

## DECISIONS (user)
- Progress is stored in Google Drive only. NO git repo. Do not suggest one.
- User said: do NOT ask for the zip (the zip is only for the user's own git upload). Work from Drive alone.

## HOW CODE GETS INTO THE SANDBOX WITHOUT A ZIP (learned in v34 chat)
- read_file_content on .ts files is LOSSY: it escapes characters (backslashes before _ * [ ] etc., \u003e for >, \u003c for <). Fine for reading and for writing tests against the APIs; NOT usable to copy code into the sandbox.
- download_file_content returns exact base64. It is the only exact route, but each file must then be re-typed into the sandbox. apply.ts alone is 26122 bytes and imports ~12 modules (~190 KB, plus xml.ts, paragraphs.ts, structure_rules.ts, version.ts underneath). Not done yet: to typecheck/run tests the sandbox needs these files. Do it file by file, check the byte size against Drive after each write, import leaf modules first (xml, paragraphs, structure_rules, version, types, zipio, xml_rewrite, ...). If a chat runs out of room, save progress (which files are in the sandbox) in the HANDOFF.
- New files that Claude writes itself (tests, new code) can be saved to Drive directly with create_file, with no transfer problem.

## 1. DONE AND TESTED (details in TEST_STATUS_v12 to v18; do not redo)
config, types.ts, zipio.ts, xml.ts, template_extract, paragraphs, structure_rules, structure_detect, xml_rewrite, styles, restyle, package_merge, numbering_merge, validate, apply_sections, page_setup, watermark (tests existed and passed as recorded in those files).

## 2. WRITTEN BUT UNTESTED
- apply.ts (26122 bytes, id 1XT49S-VXtYRsyphgtsYA_yXGC9BmbKWK, src/core, "version 2: page setup and watermark wired in"). Desk-read in v34: one thing to check when it runs: `portraitSize(pageSetup, templateSection);` is called and its result is thrown away (check page_setup.ts: if it only validates, fine; if it returns something apply needs, fix).
- apply_extras.test.ts (NEW, 16950 bytes, id 1vpPtQ93YZb6UxiJvPdwftkugiGpd7e04, src/core). 16 tests: page setup (default keeps sizes, A4 to Letter, landscape, portrait, margins, selected sections [0], custom size, validatePackage clean) and watermark (inherit, remove, text replaces, text adds, keepSource warning, no-header warning, image mode rels/media/content type once + validate clean, image errors). Its expected values are inferred from types.ts comments and apply.ts: when a test fails, decide whether the test or the code is wrong. It builds its own small zips; it imports setWatermark, validatePackage, scanXmlString, openDocx, readEntryText, setWorkersEnabled, parseRelationships.
- apply.test.ts on Drive is still the old 22087-byte, 21-test file (kept: apply_extras.test.ts is separate, so no rename needed; item 2 of REMAINS is "run it").
- NOTE: HANDOFF_v32/v33 said "validateDocx"; no such function exists. The real exports are validatePackage, scanXmlString, createPartScanner, hasErrors.

## 3. REMAINS (in this order)
1. Get the sandbox code in (see HOW CODE GETS IN), then typecheck apply.ts with apply.test.ts, apply_sections tests and apply_extras.test.ts; fix compile errors first; do not rewrite if it passes.
2. Run apply_extras.test.ts; fix test or code as needed (save changed code as a new file with the real name, old one renamed name_OLD_vN and moved to "wrong").
3. Make pageSetup, watermark, keepSourceHeaders required in FormatOptions only if tests allow.
4. 3c partial apply; 5a header/footer choices and keepSourceHeaders; 4a-4b captions and cross-references; toc.ts; rest of the feature register (v21 section 6).
5. Run the whole Drive suite once (Node 22 and 18, typecheck, build).
6. Then manual_mark.ts, web UI, template_store.ts, report.ts, image_reduce.ts, batch.ts, deploy guide.

## 4. RESUME PROMPT (give this at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Progress is kept in Drive only (no git, no zip needed). If I attach FORMAT_IQ.zip, unzip it into /home/claude/formatiq; if not, do not ask me for it: bring the code in from Drive as described in the latest HANDOFF.

START: open the highest-numbered HANDOFF_vN.md in Drive and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Drive wins over any zip if they differ.

Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER.

Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested.

After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders. Every new HANDOFF must contain, inside itself (not just a pointer to older files): (1) what is DONE and tested, (2) what is written but UNTESTED, (3) the full list of what REMAINS, in order, (4) this resume prompt.

END: give me a short message saying what was built, tested (with numbers) and what remains, and give me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save or test.
