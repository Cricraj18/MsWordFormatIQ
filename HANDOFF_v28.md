# formatIQ HANDOFF v28 (short update on top of HANDOFF_v21, v23, v24, v25, v26 and v27)

READ ORDER: (1) HANDOFF_v21.md. (2) v23. (3) v24. (4) v25. (5) v26. (6) v27 (types.ts v2, sandbox rebuild tips). (7) This file. (8) TEST_STATUS_v12, v13, v14, v15, v16 (each is an update on top of the one before). (REVIEW and progress files: optional history.)
Trust order if files disagree: v28, v27, v26, v25, v24, v23, v21, TEST_STATUS_v16, v15, v14, v13, v12, REVIEW.
CLEAN-UP RULE: as in v21 section 8, with one change: keep HANDOFF_v21, v23, v24, v25, v26, v27 AND v28 in place and keep TEST_STATUS_v12 to v16 in place. Everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong".
CONSOLIDATION: still not done (seven handoffs, five test statuses). Needs a chat with plenty of room: ONE self-contained HANDOFF_v29 and ONE TEST_STATUS_v17, then move the old ones to "wrong".
TWO CHATS AT ONCE: before saving a file or a handoff, search Drive for the same name or a higher HANDOFF number; if one exists, read it and continue from it.

## STATUS (chat 16)
Clean-up search of all four folders (root, src, src/core, wrong): nothing needed moving.
DONE, saved in Drive, tested in sandbox (Node 22.22.2 AND Node 18.20.8, typecheck clean, build ok):
- src/core/page_setup.ts VERSION 2, 16360 bytes (id 12TRoyShP-ll8qLr9mn9TZrS4vg5Jt6Ij). It imports PagePreset, Orientation, MarginSet, PageSetupOptions from "../types" and re-exports them (export type {...}). Nothing else changed. The unchanged page_setup.test.ts (16895 bytes, id 194pt_PIviDzLZzjiGfeK71fPaq6SVxjW) ran against it: 39 of 39 pass on both Node versions. The old 16829-byte file is in "wrong" as page_setup_OLD_v27.ts (id 1CFtvu_YW5C-LbERbsnORct3YjLwrR6v7).
Note: update_file answered "No approval received" twice at first, then worked on the third try. Retry before giving up.

## NEXT (in this order)
1. watermark.ts the same way: import WatermarkMode and WatermarkOptions from "../types", re-export, run its 35 tests (watermark.test.ts 14311 bytes), save as a new file, old one to "wrong" as watermark_OLD_v28.ts. Needs xml_rewrite.ts, package_merge.ts, types.ts loaded in the sandbox.
2. Make the three FormatOptions fields required only if apply.test.ts allows (load and typecheck it first).
3. INTEGRATE page_setup.ts and watermark.ts into apply.ts (v25 NEXT 1, v26 HOW TO WIRE); tests in apply.test.ts; replace apply.ts and apply.test.ts the safe way (OLD_v28 names).
4. 3c partial apply, 5a header/footer choices and keepSourceHeaders, 4a-4b captions + cross-references, toc.ts, rest of the feature register.
5. Run validateDocx on an apply output with watermark and page setup; run the whole Drive suite once.

## SANDBOX NOTES (chat 16)
Sandbox was reset again. Rebuilt in /home/claude/formatiq (package.json, tsconfig, vite.config.ts), Node 18 at /tmp/n18. Loaded only: comment-free types.ts (same exports as Drive v2), comment-free xml_rewrite.ts, page_setup.ts v2, page_setup.test.ts. The sandbox cannot reach Drive by network; read_file_content works for code (unescape backslashes before _ [ ] * ! < >), then type the file into the sandbox. Build check: temporary src/main.ts and index.html, npm run build, delete them.

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER. Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested. After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
