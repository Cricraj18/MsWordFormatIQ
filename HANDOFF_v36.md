# formatIQ HANDOFF v36 (self-contained; short update on top of HANDOFF_v21, v23 to v35)

READ ORDER: (1) HANDOFF_v21. (2) v23 to v35. (3) This file. (4) TEST_STATUS_v12 to v23. (REVIEW and progress files: optional history.)
Trust order if files disagree: v36, v35, v34, ... v23, v21, TEST_STATUS_v23 down to v12, REVIEW.
CLEAN-UP RULE: keep HANDOFF_v21 and v23 to v36, TEST_STATUS_v12 to v23, highest progress file (progress_v28). Everything else older goes to "wrong" (Drive ID 1306nkF0ShJox257WewKZ2s6QK51fzgda). Never delete.

## CODE LOCATION (changed in v36)
Code now lives in the git repo https://github.com/Cricraj18/MsWordFormatIQ (user pushes it; the sandbox cannot push, no credentials). Start every chat by cloning that repo into /home/claude/formatiq (sandbox can reach github.com), or unzip FORMAT_IQ.zip if attached. The old "Drive only, no git" decision is replaced by: code in git, handoffs also in Drive. Drive code copies of types.ts, types.test.ts and apply.test.ts are OLDER than the repo (see section 2); the repo/zip wins.
At the time of v36 the repo was empty or not readable from the sandbox (git ls-remote printed nothing); the code came from the user's FORMAT_IQ zip.

## 1. DONE AND TESTED (v36 chat: typecheck clean, Node 22 and Node 18: 19 files, 417 tests all passed, build ok)
config, types.ts, zipio, xml, template_extract, paragraphs, structure_rules, structure_detect, xml_rewrite, styles, restyle, package_merge, numbering_merge, validate, apply_sections, page_setup (39 tests), watermark (35), apply.ts (apply.test.ts 21 tests + apply_extras.test.ts 16 tests: page setup, watermark, image watermark, validate clean) .
Per file counts: apply 21, apply_extras 16, page_setup 39, zipio 28, package_merge 36, watermark 35, validate 21, template_extract 27, restyle 21, numbering_merge 29, structure_detect 14, structure_rules 24, xml_rewrite 32, paragraphs 18, styles 22, apply_sections 11, xml 12, types 9, version 2 = 417.
The portraitSize(...) call in apply.ts is an early validation that throws on bad custom sizes; result unused on purpose. Fine.
FormatOptions.pageSetup, watermark and keepSourceHeaders are now REQUIRED (types.ts). Tests changed: apply.test.ts line 3 and the "selected" test use createDefaultFormatOptions(); types.test.ts old "hand-made options" test replaced by one that checks the defaults.

## 2. CHANGED BUT NOT YET IN DRIVE (repo/zip has them; Drive copies are the old ones)
src/types.ts (3 fields required), src/types.test.ts, src/core/apply.test.ts. Next chat: if working from the repo, nothing to do. If working from Drive, replace these three the safe way.

## 3. WRITTEN BUT UNTESTED
Nothing in src. (Not tested anywhere: real browser, real Word/LibreOffice rendering, 1 GB file.)

## 4. REMAINS (in order)
1. 3c partial apply (scope "selected", FormatPart list: styles, headersFooters, toc, crossReferences, sections). apply.ts currently throws INVALID_OPTIONS for "selected".
2. 5a header/footer choices and keepSourceHeaders (apply only warns for watermark keepSource; source headers are dropped).
3. 4a-4b captions and cross-references; toc.ts.
4. Rest of the feature register (HANDOFF_v21 section 6: tables, images, lists, footnotes, page numbering, cover page, track changes and comments policy, document properties, pass-through parts).
5. manual_mark.ts, web UI, template_store.ts, report.ts, image_reduce.ts, batch.ts, deploy guide.
6. Re-run the whole suite (Node 22 and 18, typecheck, build) after each step.

## 5. RESUME PROMPT (give this at the end of every chat, in a code block)
My code is in this git repo: https://github.com/Cricraj18/MsWordFormatIQ . Clone it into /home/claude/formatiq first (the sandbox can reach github.com). If I attach FORMAT_IQ.zip, unzip it instead; do not ask me for a zip otherwise. Drive folder (handoffs and history): https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Repo wins over Drive for code.

START: open the highest-numbered HANDOFF_vN.md (repo or Drive) and read it fully, then every file its READ ORDER names. Follow its CLEAN-UP RULE first (update_file moves and renames Drive files; "wrong" folder ID 1306nkF0ShJox257WewKZ2s6QK51fzgda; never delete). Then follow its WORK ORDER.

Rules: do all testing yourself in the sandbox (typecheck, tests on Node 22 and Node 18, build); never ask me to test or run anything; do not stop at test gates, keep building until your limit. When a code file must change, test the fixed version first, then save it under the real name. Do not redo anything marked tested.

After each file: save the file, the next HANDOFF_v(N+1).md, progress and TEST_STATUS files to Drive, and give me a zip of the repo folder to push to git (you cannot push). Every new HANDOFF must contain (1) DONE and tested, (2) written but UNTESTED, (3) full list of what REMAINS in order, (4) this resume prompt.

END: short message with what was built, tested (with numbers), what remains, and this same prompt in a code block. If you cannot write somewhere, say so plainly and give me the full file. Do not pretend to save or test.
