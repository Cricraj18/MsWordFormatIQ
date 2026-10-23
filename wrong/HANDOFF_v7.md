# formatIQ HANDOFF v7 (short update on top of HANDOFF_v6.md)

READ ORDER: (1) HANDOFF_v6.md fully (all rules, spec, DONE list, file names: it is still valid; do NOT move it to "wrong" yet). (2) This file. (3) REVIEW_v1.md (review findings).
If this file and v6 disagree, this file wins.

## What happened in chat 6
- 0c/0d zipio.ts, zipio.test.ts and types.ts were checked by reading (REVIEW_v1.md). No type errors, no wrong zip.js calls found. 3 medium, 7 low, 5 test points. Nothing was changed in the code yet.
- All code files in Drive are complete. Old versions are already in "wrong".
- Last chat stopped before saving progress_v5/v6. progress_v6.md is saved now.
- HANDOFF_v5.md and progress_v4.md are still in the project root. They are outdated. Drive tools here cannot move files, so ask the user to drag them into "wrong" (one line, then continue).

## NEXT STEP A (do first, small): apply the review fixes
1. Rebuild the project in the sandbox (see HOW CLAUDE CAN TEST in v6). Use download_file_content for code.
2. Apply, in this order: M3 (types.ts factory createDefaultFormatOptions; check nothing else uses DEFAULT_FORMAT_OPTIONS), M2 (control param in openDocx), M1 (size check in copyEntry), L1, L2, L4, L5, T1 to T4. Skip L3, L6, L7 (decisions for later; write them in progress notes).
3. Add tests for: openDocx with an aborted control; isSafeZipPath("a//b") is false; createDefaultFormatOptions returns a new parts array each time.
4. Run npm run typecheck, npm test, npm run build. If Node 18 is not available, say "Node 18 not tested" (try: npx -y node@18 to run the tests; UNSURE if the network allows it).
5. Save as NEW files: zipio.ts, zipio.test.ts, types.ts in their real folders. The old copies must go to "wrong" with names zipio_OLD_v2.ts etc. If the move fails, tell the user and continue. Update the export list in the next HANDOFF.
6. Save HANDOFF_v8.md (full text: v6 body with this step marked DONE) and progress_v7.md. Move HANDOFF_v6, HANDOFF_v7 and old progress files to "wrong".

## NEXT STEP B: file 1a (template_extract.ts) then 1b, exactly as in HANDOFF_v6 "NEXT FILE". Only after step A passes.

## Rules that still apply
- No "any", no "!" non-null assertion. Short messages (user reads on a phone). Say "tested" only for what was run. Stop at TEST GATE.
- If a file in Drive is cut off, empty, or does not match its handoff line count or export list: treat it as corrupt. Rename the copy with _OLD_vN, move it to "wrong", rebuild only that file, and start again from the step that needs it.

## RESUME PROMPT (give this exact text at the end of every chat, in a code block)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully. If it says to read an older HANDOFF too (HANDOFF_v7 says to read HANDOFF_v6 first), read both, plus REVIEW_v1.md. Follow them exactly: do NEXT STEP A before anything else, then NEXT STEP B. Do not redo anything in DONE. First check that every code file in Drive is complete; any incomplete, corrupt or wrong file goes to the folder "wrong" (renamed _OLD_vN) and is rebuilt from where it is needed. Save new files into the right folders, save the next HANDOFF_v(N+1).md and progress_v(N+1).md, and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
