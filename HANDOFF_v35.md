# formatIQ HANDOFF v35 (self-contained; short update on top of HANDOFF_v21, v23 to v34)

READ ORDER: (1) HANDOFF_v21. (2) v23 to v34. (3) This file. (4) TEST_STATUS_v12 to v22. (REVIEW and progress files: optional history.)
Trust order if files disagree: v35, v34, v33, ... v23, v21, TEST_STATUS_v22 down to v12, REVIEW.
CLEAN-UP RULE: keep HANDOFF_v21 and v23 to v35 in place, keep TEST_STATUS_v12 to v22 in place, keep the highest progress file (progress_v27). Everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda). Never delete.
Clean-up done in v35 chat: nothing needed moving (all files match the keep list; REVIEW_v1 and addendum kept as history).

## DECISIONS (user)
- Progress is stored in Google Drive only. NO git repo. Do not suggest one.
- User said: do NOT ask for the zip. Work from Drive alone.

## WHAT HAPPENED IN v35 CHAT
No zip, sandbox empty (Node 22.22 only; no /home/claude/formatiq). Read HANDOFF v34, v33, v21 and TEST_STATUS_v21, listed all Drive folders (src: types.ts 11519, version.ts, version.test.ts, types.test.ts; src/core: 29 files, all as listed in v34 and v21). Re-read apply.ts (26122 bytes) via read_file_content to size the import job. NOTHING was typechecked, tested or built. No code changed.
SIZE OF THE IMPORT JOB (measured): apply.ts alone is about 12K tokens to read; the chain it needs (zipio 14224, xml 4065, types 11519, template_extract 6962, paragraphs 10805, structure_rules 8254, structure_detect 4925, xml_rewrite 8447, styles 6176, restyle 7605, package_merge 14260, numbering_merge 13620, validate 11565, apply_sections 10658, page_setup 16360, watermark 17909, version 284) is about 190 KB, plus tests apply.test.ts 22087, apply_extras.test.ts 16950, apply_sections.test.ts 5442 (about 45 KB). Reading is about 70K tokens and re-typing the same again. The sandbox cannot reach Drive, so there is no copy tool. Plan if it must be done without a zip: one chat does ONLY the import (leaf first: version, types, xml, zipio, xml_rewrite, paragraphs, structure_rules, structure_detect, template_extract, styles, restyle, package_merge, numbering_merge, apply_sections, validate, page_setup, watermark, then apply.ts, then the three test files); strip comments in sandbox copies to save tokens; use read_file_content and undo the escapes (backslash before _ * [ ] ! and the codes for less-than and greater-than; a doubled backslash in the view is one real backslash); no byte-size check is possible after stripping comments, so compare by typecheck + the recorded test counts instead. Record in the next HANDOFF which files are already in the sandbox (it is wiped when the chat ends, so import and test in the SAME chat).
Cheapest route for the user: attach FORMAT_IQ.zip (one download of the Drive folder). It removes the whole import cost. User said not to ask; this is only a note on cost.

## 1. DONE AND TESTED (details in TEST_STATUS_v12 to v18; do not redo)
config, types.ts, zipio.ts, xml.ts, template_extract, paragraphs, structure_rules, structure_detect, xml_rewrite, styles, restyle, package_merge, numbering_merge, validate, apply_sections, page_setup, watermark (tests existed and passed as recorded in those files).

## 2. WRITTEN BUT UNTESTED
- apply.ts (26122 bytes, id 1XT49S-VXtYRsyphgtsYA_yXGC9BmbKWK, src/core, "version 2: page setup and watermark wired in"). Check when it runs: `portraitSize(pageSetup, templateSection);` is called and its result is thrown away (check page_setup.ts: if it only validates, fine; if it returns something apply needs, fix). Second note from re-read: for watermark mode "keepSource" apply only warns and uses the template's headers (by design for now).
- apply_extras.test.ts (16950 bytes, id 1vpPtQ93YZb6UxiJvPdwftkugiGpd7e04, src/core): 16 tests (page setup 8, watermark 8). Expected values inferred from types.ts comments and apply.ts: when a test fails, decide whether test or code is wrong. Imports setWatermark, validatePackage, scanXmlString, openDocx, readEntryText, setWorkersEnabled, parseRelationships.
- apply.test.ts on Drive is the old 22087-byte, 21-test file (id 17zH_WNu75sRej8ls4pEqTGWcIxl0qRcD), kept.
- NOTE: earlier handoffs said "validateDocx"; no such function exists. Real exports: validatePackage, scanXmlString, createPartScanner, hasErrors.

## 3. REMAINS (in this order)
1. Get the sandbox code in (see WHAT HAPPENED: zip, or an import-only chat), then typecheck apply.ts with apply.test.ts, apply_sections tests and apply_extras.test.ts; fix compile errors first; do not rewrite if it passes.
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
