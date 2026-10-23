# formatIQ HANDOFF v9 (short update; replaces HANDOFF_v8.md and v7)

READ ORDER: (1) HANDOFF_v6.md fully (rules, spec, DONE list, export names). (2) HANDOFF_v8.md (corrections, step A fix list). (3) This file. (4) TEST_STATUS_v1.md. (5) REVIEW_v1.md.
If files disagree, trust in this order: v9, v8, TEST_STATUS_v1, REVIEW files, v6.

## NEW WORKING RULES (the user asked for these; they override v6 and v8)
1. NO MANUAL TESTING BY THE USER. Claude tests everything itself in the sandbox (bash_tool): typecheck, tests, build, and probe scripts. Never ask the user to run npm, open a browser, or go to the office.
2. NO TEST GATE STOPS. Do not stop and wait. Keep building the next files in FILE LIST order, one after the other, until the chat limit. After each file: write it, write its tests, run typecheck + tests (Node 22 AND Node 18) + build, then save to Drive.
3. Test whenever a file is new or changed. If a test fails, fix the code or the test and run again. Say "tested in sandbox" only for what really ran, with numbers.
4. Things the sandbox cannot test (real browser Web Workers, file picker, showSaveFilePicker, dev server page): first TRY once to get a headless browser (npm i playwright or puppeteer; check x-deny-reason if the network blocks the browser download; apt is allowed only from archive.ubuntu.com). If it cannot work, write "browser: not testable here" in TEST_STATUS, keep that code small and defensive, and CONTINUE. Do not ask the user.
5. Make a NEW file whenever a file must change (Drive tools cannot overwrite, move or rename): save the new version with the real name, and the old version as name_OLD_vN in the same folder. Tell the user ONE line: "Please move these old files into 'wrong': ...". Never wait for it.
6. Keep all other rules from v6: no "any", no "!" non-null assertion, short messages for a phone, complete files, doc comments, say UNSURE only if it could not be tested.

## WORK ORDER
A. Apply the fix list in HANDOFF_v8.md "NEXT STEP A" (M3, corrected M2, L1, L2, L4, L5, T2, T3, T4 and the 4 new tests). Test on Node 22 and Node 18 + build.
B. Sandbox-testable items from TEST_STATUS_v1.md "NOT TESTED": 1 GB-class file (try about 900 MB; if RAM is short, test 600 MB again and stop), zip.js passThrough, a real-style .docx (build one with a library from npm, for example the "docx" package, and open it with openDocx), headless browser attempt (rule 4).
C. Then build in order: 1a template_extract.ts, 1b its tests, 2a structure_detect.ts, 2b tests, 3a-3c, and so on (HANDOFF_v6 FILE LIST). Choose an XML parser that works in browser AND Node and does not expand entities; test it with a malicious entity sample.
D. After EVERY file or two: save the file(s) to the right Drive folder, then save HANDOFF_v(N+1).md (short update, same READ ORDER plus this file) and progress_v(N+1).md, and update TEST_STATUS as TEST_STATUS_v(N+1).md (what is tested, what is not).

## FILE CHECK AT THE START OF EVERY CHAT
List the Drive folders (FORMAT_IQ root, src, src/core, wrong). Download each code file (download_file_content, base64, replace the text "\u003d" with "=") and compare its size and export list with the handoff. If a file is cut off, empty, or does not typecheck: save a copy named name_OLD_vN, tell the user one line to move it to "wrong", rebuild only that file and start again from the step that needs it. Rebuild the project in /home/claude/formatiq first and run typecheck + tests before any new work.

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully, then every file its READ ORDER names (older HANDOFF files, TEST_STATUS, REVIEW). Follow its WORK ORDER. Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building the next files until your limit. First check every code file in Drive is complete; if one is incomplete, corrupt or wrong, save a fixed new file and the old one as name_OLD_vN, tell me in one line to move the old one into the folder "wrong", and continue from where it is needed. Do not redo anything marked tested. After each file save the file, the next HANDOFF_v(N+1).md, progress_v(N+1).md and TEST_STATUS_v(N+1).md into the right folders, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
