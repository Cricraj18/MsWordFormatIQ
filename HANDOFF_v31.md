# formatIQ HANDOFF v31 (short update on top of HANDOFF_v21, v23 to v30)

READ ORDER: (1) HANDOFF_v21. (2) v23 to v29. (3) v30. (4) This file. (5) TEST_STATUS_v12 to v19. (REVIEW and progress files: optional history.)
Trust order if files disagree: v31, v30, v29, v28, v27, v26, v25, v24, v23, v21, TEST_STATUS_v19 down to v12, REVIEW.
CLEAN-UP RULE: as in v30, keep HANDOFF_v21, v23 to v31 and TEST_STATUS_v12 to v19 in place; everything else older than the highest HANDOFF/progress/TEST_STATUS goes to "wrong".

## STATUS (chat 20, short: no code work was possible)
- Drive opened and written fine. Searched root, src, src/core: nothing needed moving.
- In this chat I fully read v30 and v21 only (v23 to v29 and TEST_STATUS v12 to v18 were NOT re-read; v30 says it read them). 
- The sandbox was empty again and has NO network path to Drive, so apply.ts (26122 bytes, id 1XT49S-VXtYRsyphgtsYA_yXGC9BmbKWK) is still UNTESTED. Nothing was changed in code.
- Lesson: code must be moved into the sandbox by download_file_content (base64), which costs a lot of chat room. Start the next chat with the sandbox rebuild immediately, before reading anything else twice.

## NEXT
Same list as HANDOFF_v30 section NEXT (1 to 5): rebuild the sandbox for apply.ts, typecheck it with the old apply.test.ts and apply_sections tests, add the new tests, then 3c, 5a, 4a/4b, toc.ts, whole suite.

## RESUME PROMPT
Same text as in HANDOFF_v30 / v21 section 12. Give it at the end of every chat in a code block.
