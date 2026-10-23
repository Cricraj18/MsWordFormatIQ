# formatIQ HANDOFF v10 (short update on top of v9; fixes the "cannot move files" mistake)

READ ORDER: (1) HANDOFF_v6.md (rules, spec, DONE list, export names). (2) HANDOFF_v8.md (corrections, step A fix list). (3) HANDOFF_v9.md (working rules, WORK ORDER). (4) This file. (5) TEST_STATUS_v1.md (or the highest TEST_STATUS_vN). (6) REVIEW_v1.md.
Trust order if files disagree: v10, v9, v8, TEST_STATUS, REVIEW, v6.
Keep HANDOFF_v6, v8, v9 in place (they are read each time) until a later handoff copies their content in full.

## CORRECTION: Claude CAN move and rename Drive files
Chat 6 wrongly said Drive tools cannot move files. They can. Tool: mcp__Google_Drive__update_file (load it with tool_search first: "Google Drive update_file move rename"). It takes fileId plus parentId (moves the file) and/or title (renames it). It cannot change content; for new content create a new file (create_file) and keep the old one as name_OLD_vN.
Folder IDs: project FORMAT_IQ 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA; src 108HasgSfuacTIwGOV-ugUoqXmMp6nkCl; src/core 1gD5Q2WEWwISUiOoFhZ-kblS0pMPjlIxf; wrong 1306nkF0ShJox257WewKZ2s6QK51fzgda.
Done in chat 6: moved HANDOFF_v5.md, progress_v4.md, HANDOFF_v7.md into "wrong".

## CLEAN-UP RULE (do at the START of every chat and again after every save)
1. Search each folder (search_files with parentId = folder id; excludeContentSnippets true).
2. A file is BAD if: it is outdated (an older HANDOFF_vN, progress_vN or TEST_STATUS_vN than the highest, except the files named as "keep" above); it is incomplete (cut off, empty, size far below its handoff line count or export list, or fails typecheck/tests); it is corrupt or wrong; or its name has _OLD or "ignore" in it but it sits outside "wrong".
3. MOVE every BAD file into "wrong" with update_file (fileId, parentId = 1306nkF0ShJox257WewKZ2s6QK51fzgda). If two files in "wrong" have the same name, rename the older with update_file title (add _OLD_vN) first. NEVER trash or delete (trash_file is not allowed).
4. When a code file must be replaced: first rebuild the fixed version and test it in the sandbox; create the new file with create_file under the real name in the real folder; then rename the OLD file (update_file title = name_OLD_vN) and move it to "wrong". Do it in this order so the project is never left without a working file. If the old and new have the same name, rename the old one first, then create the new one.
5. Repeat the search to check the folders: only good, current files remain outside "wrong". Report in ONE short line what was moved. If update_file fails (needs approval or an error), say so in one line, and continue. Do not ask the user to move files unless update_file really failed twice.
6. Keep root clean: root should hold only the highest HANDOFF/progress/TEST_STATUS files, the kept HANDOFF_v6/v8/v9, REVIEW files, project config files, and the src folder.

## Everything else
Follow HANDOFF_v9.md WORK ORDER and NEW WORKING RULES (all testing by Claude in the sandbox, no test gates, no manual testing by the user). After each file or two, save the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md. When a new handoff is saved, it should contain this CLEAN-UP RULE and the folder IDs (copy them in).

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Follow its CLEAN-UP RULE first: you CAN move and rename Drive files with the update_file tool (load it with tool_search; parentId moves a file, title renames it). Move every outdated, incomplete, corrupt or wrong file into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda) yourself; never delete; do not ask me to move anything unless update_file fails twice. Then follow its WORK ORDER. Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. When a code file must change, test the fixed version first, save it as a new file with the real name, and rename and move the old one to "wrong" as name_OLD_vN. Do not redo anything marked tested. After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
