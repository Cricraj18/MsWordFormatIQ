# formatIQ HANDOFF v11 (short update on top of v6, v8, v9; includes everything from v10)

READ ORDER: (1) HANDOFF_v6.md (rules, spec, DONE list, export names). (2) HANDOFF_v8.md (corrections; its step A is now DONE). (3) HANDOFF_v9.md (NEW WORKING RULES, WORK ORDER). (4) This file. (5) TEST_STATUS_v2.md (or the highest TEST_STATUS_vN). (6) REVIEW_v1.md and REVIEW_v1_addendum.md.
Trust order if files disagree: v11, v9, v8, TEST_STATUS_v2, REVIEW, v6.
Keep HANDOFF_v6, v8, v9 in place (read each time) until a later handoff copies their content in full.

## STATUS (chat 7)
- Step A DONE and saved: types.ts, zipio.ts, zipio.test.ts. Typecheck clean, 24/24 tests on Node 22.22.2 and Node 18.20.8, build ok (see TEST_STATUS_v2).
- Name/API changes to reuse: types.ts now exports createDefaultFormatOptions() (a function; DEFAULT_FORMAT_OPTIONS no longer exists). openDocx(source, fileName, limits?, control?) has a 4th param. isSafeZipPath rejects empty segments ("a//b") and allows one trailing "/". Everything else in HANDOFF_v6 DONE list is unchanged.
- Old copies of the three files are in "wrong" as types_OLD_v2.ts, zipio_OLD_v2.ts, zipio.test_OLD_v2.ts.
- NEXT: WORK ORDER step B (HANDOFF_v9), then C: 1a template_extract.ts, 1b tests, 2a structure_detect.ts, and so on in FILE LIST order (HANDOFF_v6).

## HOW TO REBUILD THE PROJECT IN THE SANDBOX
The sandbox cannot reach Drive by network. Read each code file with read_file_content (clean text with markdown backslashes before [ ] * _ ! < > and "\u003d" style escapes; remove those) or download_file_content (base64), then write the files under /home/claude/formatiq with bash/create_file, npm install, and run typecheck + tests. Check sizes against Drive file sizes (types.ts 7947, zipio.ts 13382, zipio.test.ts 10874 bytes). Node 18: see TEST_STATUS_v2.

## DRIVE WRITE APPROVAL (learned in chat 7)
Drive write tools (update_file, create_file) first answered "No approval received" until the user approved the permission prompt in the Claude app. If it happens: try once more (this shows the prompt), then tell the user plainly and wait for the approval; never pretend to save.

## FOLDER IDS
Project FORMAT_IQ 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA; src 108HasgSfuacTIwGOV-ugUoqXmMp6nkCl; src/core 1gD5Q2WEWwISUiOoFhZ-kblS0pMPjlIxf; wrong 1306nkF0ShJox257WewKZ2s6QK51fzgda.

## CLEAN-UP RULE (do at the START of every chat and again after every save)
1. Search each folder (search_files with parentId = folder id; excludeContentSnippets true; follow nextPageToken).
2. A file is BAD if: it is outdated (an older HANDOFF_vN, progress_vN or TEST_STATUS_vN than the highest, except the keep files HANDOFF_v6, v8, v9); it is incomplete (cut off, empty, size far below its expected size or export list, or fails typecheck/tests); it is corrupt or wrong; or its name has _OLD or "ignore" in it but sits outside "wrong".
3. MOVE every BAD file into "wrong" with update_file (fileId, parentId = 1306nkF0ShJox257WewKZ2s6QK51fzgda). If two files in "wrong" would have the same name, rename the older with update_file title (add _OLD_vN) first. NEVER trash or delete.
4. When a code file must be replaced: first rebuild the fixed version and test it in the sandbox; then rename the OLD file (update_file title = name_OLD_vN) and move it to "wrong"; then create the new file with create_file (textContent, contentMimeType text/plain, disableConversionToGoogleType true) under the real name in the real folder. Compare the new file size with the sandbox size.
5. Repeat the search to check. Report in ONE short line what was moved. If update_file fails twice, say so in one line and ask the user to approve or move by hand.
6. Keep root clean: root holds only the highest HANDOFF/progress/TEST_STATUS files, kept HANDOFF_v6/v8/v9, REVIEW files, config files (package.json, tsconfig.json, vite.config.ts, index.html), the src folder and the "wrong" folder.

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER. Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested. After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
