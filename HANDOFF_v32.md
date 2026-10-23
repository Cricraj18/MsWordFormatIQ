# formatIQ HANDOFF v32 (short update on top of HANDOFF_v21, v23 to v31)

READ ORDER: (1) HANDOFF_v21. (2) v23 to v31. (3) This file. (4) TEST_STATUS_v12 to v19. (REVIEW and progress files: optional history.)
Trust order if files disagree: v32, v31, v30, v29, v28, v27, v26, v25, v24, v23, v21, TEST_STATUS_v19 down to v12, REVIEW.
CLEAN-UP RULE: keep HANDOFF_v21 and v23 to v32 in place, keep TEST_STATUS_v12 to v19 in place, keep the highest progress file. Everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong".

## WHY THE SANDBOX IS ALWAYS EMPTY
The sandbox is wiped when a chat ends, and it cannot reach Drive by network (only npm, pypi, github). Code can only get in by (a) a zip the user attaches, or (b) download_file_content (base64, very costly). From now on the user attaches FORMAT_IQ.zip (download of the Drive folder) at the start of each chat. Unzip it into /home/claude/formatiq first. If no zip is attached, say so in one line and use (b) only for the files the next step imports.

## DONE AND TESTED (details in TEST_STATUS_v12 to v18; do not redo)
config, types.ts, zipio.ts, xml.ts, template_extract, paragraphs, structure_rules, structure_detect, xml_rewrite, styles, restyle, package_merge, numbering_merge, validate, apply_sections, page_setup, watermark (tests existed and passed as recorded in those files).

## WRITTEN BUT UNTESTED
apply.ts (26122 bytes, id 1XT49S-VXtYRsyphgtsYA_yXGC9BmbKWK, "version 2: page setup and watermark wired in"). apply.test.ts on Drive is still the old 22087-byte, 21-test file.

## REMAINS (in this order)
1. Typecheck the new apply.ts with the old apply.test.ts and apply_sections tests; fix compile errors first; do not rewrite if it passes.
2. Add new tests: A4 to Letter, landscape, margins, source landscape kept, selected sections; template with watermark header and each mode; no header plus "text" gives the warning; image mode adds rels, media and content type exactly once; validateDocx on an output has no errors. Replace apply.test.ts the safe way (old one renamed apply.test_OLD_vN.ts and moved to "wrong").
3. Make pageSetup, watermark, keepSourceHeaders required in FormatOptions only if tests allow.
4. 3c partial apply; 5a header/footer choices and keepSourceHeaders; 4a-4b captions and cross-references; toc.ts; rest of the feature register (v21 section 6).
5. Run the whole Drive suite once (Node 22 and 18, typecheck, build).
6. Then manual_mark.ts, web UI, template_store.ts, report.ts, image_reduce.ts, batch.ts, deploy guide.

## RESUME PROMPT (give this at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). I have attached FORMAT_IQ.zip (a fresh download of that folder). If no zip is attached, tell me in one line and continue with the slow method (download_file_content) only for the files the next step imports.

START: unzip the attachment into /home/claude/formatiq first. The sandbox is wiped between chats, so this zip is the only way code gets in. Then open the highest-numbered HANDOFF_vN.md in Drive (or in the zip) and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Compare the zip's files with Drive (names and sizes); Drive wins if they differ.

Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER.

Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested.

After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders. Every new HANDOFF must contain, inside itself (not just a pointer to older files): (1) what is DONE and tested, (2) what is written but UNTESTED, (3) the full list of what REMAINS, in order, (4) this resume prompt.

END: give me a short message saying what was built, tested (with numbers) and what remains, tell me to download the FORMAT_IQ folder as a zip again for the next chat, and give me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save or test.
