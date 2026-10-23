# formatIQ HANDOFF v29 (short update on top of HANDOFF_v21, v23, v24, v25, v26, v27 and v28)

READ ORDER: (1) HANDOFF_v21.md. (2) v23. (3) v24. (4) v25. (5) v26. (6) v27. (7) v28. (8) This file. (9) TEST_STATUS_v12, v13, v14, v15, v16, v17 (each is an update on top of the one before). (REVIEW and progress files: optional history.)
Trust order if files disagree: v29, v28, v27, v26, v25, v24, v23, v21, TEST_STATUS_v17, v16, v15, v14, v13, v12, REVIEW.
CLEAN-UP RULE: as in v21 section 8, with one change: keep HANDOFF_v21, v23, v24, v25, v26, v27, v28 AND v29 in place and keep TEST_STATUS_v12 to v17 in place. Everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong".
CONSOLIDATION: still not done (eight handoffs, six test statuses). Needs a chat with plenty of room: ONE self-contained HANDOFF_v30 and ONE TEST_STATUS_v18, then move the old ones to "wrong".
TWO CHATS AT ONCE: it happened again in chat 17: another chat saved watermark.ts v2 (13:25 UTC) and moved the old file to "wrong" while this chat was starting, and left no handoff or test record. Before saving a file or a handoff, search Drive for the same name or a higher HANDOFF number; if one exists, read it and continue from it. Also search for files modified in the last hour (search_files query modifiedTime > '...') to see whether another chat is active.

## STATUS (chat 17)
Clean-up search of all four folders (root, src, src/core, wrong): nothing needed moving except progress_v21 (older than progress_v22).
DONE, tested in sandbox (Node 22.22.2 AND Node 18.20.8, typecheck clean, build ok):
- src/core/watermark.ts VERSION 2, 17909 bytes (id 1VFty_ZiM-_bH_WLK-dl07t_Wp_b3lG9A): imports WatermarkMode, WatermarkOptions from "../types" and re-exports them. The unchanged watermark.test.ts (14311 bytes, id 1QDlxekDALkiJxkw04EhtZDiku6gXr4U4): 35 of 35 pass on both Node versions. Old file in "wrong" as watermark_OLD_v28.ts (id 1as0FPD8oSEr3EHdT2ZsBUbkGZ0vvdvl1). This is v28 NEXT 1: DONE.
Nothing else changed in Drive code.

## NEXT (in this order)
1. Make the three FormatOptions fields (pageSetup, watermark, keepSourceHeaders) required only if apply.test.ts allows (load and typecheck it first).
2. INTEGRATE page_setup.ts and watermark.ts into apply.ts (v25 NEXT 1, v26 HOW TO WIRE); tests in apply.test.ts; replace apply.ts and apply.test.ts the safe way (apply_OLD_v29.ts, apply.test_OLD_v29.ts).
3. 3c partial apply, 5a header/footer choices and keepSourceHeaders, 4a-4b captions + cross-references, toc.ts, rest of the feature register.
4. Run validateDocx on an apply output with watermark and page setup; run the whole Drive suite once.

## SANDBOX NOTES (chat 17)
Sandbox was reset again. Rebuilt in /home/claude/formatiq (package.json, tsconfig, vite.config.ts), Node 18 at /tmp/n18. Loaded: comment-free types.ts (4010 bytes), xml_rewrite.ts (5875), package_merge.ts (9562), watermark.ts (14150), watermark.test.ts. Everything was typed from read_file_content text (unescape: backslash before _ [ ] * ! < > is dropped; "\\t" and "\\n" in regexes and strings are ONE backslash in the source; "\&amp;" is "&amp;"). The sandbox cannot reach Drive by network. For apply.ts you also need apply_sections, numbering_merge, restyle, styles, zipio, paragraphs, structure_rules, structure_detect, template_extract and the tests (about 190 KB in all, v27 lists sizes): plan the chat around loading only what is needed.

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER. Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested. After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
