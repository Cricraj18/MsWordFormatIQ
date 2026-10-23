# formatIQ HANDOFF v8 (short update on top of HANDOFF_v6.md; replaces HANDOFF_v7.md)

READ ORDER: (1) HANDOFF_v6.md fully (rules, spec, DONE list, export names). Still valid. (2) This file. (3) TEST_STATUS_v1.md (what is tested, what is not). (4) REVIEW_v1.md (findings, but CORRECTED by this file and TEST_STATUS_v1.md).
If files disagree, order of trust: this file, then TEST_STATUS_v1.md, then REVIEW_v1.md / REVIEW_v1_addendum.md, then HANDOFF_v6.md. HANDOFF_v7.md is outdated (move to "wrong" when possible).

## Corrections to REVIEW_v1 (all proven by real runs, see TEST_STATUS_v1.md)
- M1: DROP. zip.js 2.23.0 already throws "Invalid uncompressed size". Add no check.
- M2: the REVIEW_v1 fix is WRONG ("signal" is not an option of getEntries; tsc error). Do this instead: openDocx gets 4th param control?: JobControl; check control.signal.aborted before and after getEntries(); catch uses mapError(err, control, "ZIP_READ_FAILED").
- T1: DROP (Node 18 passes).
- Keep: M3, L1, L2, L4, L5, T2 (as style), T3, T4. L3 optional (message is "File format is not recognized"). L6, L7: decisions for later.

## Project facts (tested, chat 6)
Typecheck clean, 17/17 tests on Node 22 AND Node 18.20.8, build ok, 600 MB docx copy ok (5.9 s, peak RSS 1.4 GB). Node 18 for tests: mkdir /tmp/n18; cd /tmp/n18; npm init -y; npm i node@18; run: /tmp/n18/node_modules/node/bin/node node_modules/vitest/vitest.mjs run. Sandbox: 4 GB RAM, 1 CPU, no "time" command (use a log file). To type-check a probe test that uses "process", leave it out of src or skip tsc for it. Delete probe files before saving.

## NEXT STEP A (first, small): apply fixes, then re-test
1. Rebuild the project in the sandbox from Drive (download_file_content; decode base64; replace the text "\u003d" with "="). Do NOT use read_file_content for code.
2. Apply: M3 (createDefaultFormatOptions() function in types.ts; keep DEFAULT_FORMAT_OPTIONS only if used), M2 (corrected), L1 (Math.min(timeoutMs, 2_147_483_647)), L2 (path.replace(/\/$/, "").split("/").every((p) => p !== "" && p !== ".." && p !== ".")), L4 (stale UNSURE comments: say "Tested with zip.js 2.23.0"), L5 (add xlsx|pptx|pdf|heic to isAlreadyCompressed), T2 (chunks.push(chunk.slice()) and keep the .map), T3 (Math.imul in the byte generator; compare bytes directly instead of the hash; rename if needed), T4 (import ZipReader, BlobReader at the top).
3. Add tests: openDocx with an already-aborted control gives CANCELLED; isSafeZipPath("a//b") is false and "word/media/" is true; createDefaultFormatOptions() returns a new parts array each call; createJobControl(2**31 + 10) is not aborted after 50 ms.
4. Run: typecheck, npm test on Node 22 AND Node 18, npm run build. Report numbers. Say "tested" only for what ran.
5. Save changed files as NEW files in their real folders (zipio.ts, zipio.test.ts, types.ts). Rename the old copies to _OLD_v2 and move them to "wrong" if the tools allow; if not, tell the user in one line.
6. Save HANDOFF_v9.md (full text, step A marked DONE) and progress_v7.md. Update TEST_STATUS (as TEST_STATUS_v2.md).

## NEXT STEP B: the NOT TESTED list in TEST_STATUS_v1.md: items 3, 4, 5 can be tried in the sandbox (try a ~900 MB docx if RAM allows; try passThrough; build one real-style .docx). Items 2 and 7 need the user in a real browser: give the user a short checklist, do not wait. Then file 1a (template_extract.ts) and 1b exactly as in HANDOFF_v6 "NEXT FILE".

## Rules that still apply
- No "any", no "!" non-null assertion. Short messages (user reads on a phone). Stop at a TEST GATE.
- If any file in Drive is cut off, empty, or does not match its export list: rename the copy _OLD_vN, move it to "wrong", rebuild only that file, and restart from the step that needs it.
- Drive tools here can create and copy files but NOT move, rename or overwrite. Moves are done by the user.

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully. Follow its READ ORDER (it names older files, TEST_STATUS and REVIEW files to read too). Do NEXT STEP A first, then NEXT STEP B. Do not redo anything marked tested in TEST_STATUS. First check that every code file in Drive is complete; any incomplete, corrupt or wrong file goes to the folder "wrong" (renamed _OLD_vN) and is rebuilt from where it is needed. Save new files into the right folders, save the next HANDOFF_v(N+1).md and progress_v(N+1).md, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
