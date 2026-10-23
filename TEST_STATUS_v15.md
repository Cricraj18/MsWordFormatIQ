# formatIQ TEST STATUS v15 (chat 15; update on top of TEST_STATUS_v12, v13, v14; read with HANDOFF_v27)

Method: sandbox (Linux, 4 GB RAM, 1 CPU), rebuilt from scratch in chat 15. Node 22.22.2 and Node 18 (/tmp/n18). Vitest 2.1.9, TypeScript 5.9.3, Vite 5.4.21.

## TESTED (really run, chat 15)
- src/types.ts v2 (11519 bytes) + src/types.test.ts (3786 bytes): 9 of 9 pass on Node 22 AND Node 18; typecheck clean (strict, noUncheckedIndexedAccess, noUnused*); build ok with a temporary entry (bundle 1.03 kB; entry and index.html removed afterwards); ASCII only (searched); no "any" type (the word appears only in comments). Drive copies have the same byte sizes as the sandbox copies.
- Coverage: old defaults unchanged; new defaults (pageSetup template/template/template/all, watermark inherit, keepSourceHeaders false); new nested objects on every call; FormatOptions built by hand WITHOUT the new fields still compiles; full PageSetupOptions, MarginSet, WatermarkOptions text and image values compile; FormatIQError, isFormatIQError, toFormatIQError and LOW_CONFIDENCE_THRESHOLD behave as before.
- Not redone: everything in TEST_STATUS_v12, v13, v14 (page_setup 39, numbering_merge 29, validate 21, apply 21, apply_sections 11, watermark 35, and earlier chats).

## CAVEATS (open; v12, v13, v14 caveats still apply)
1. The importers of types.ts were NOT typechecked together with the new file: only types.ts and its own test were in the sandbox. apply.ts and zipio.test.ts were READ (they use only createDefaultFormatOptions, format.scope, format.updateFieldsOnOpen: safe). apply.test.ts was not read; the three new FormatOptions fields are optional so a hand-made FormatOptions there still compiles. Run the whole Drive suite and typecheck once.
2. page_setup.ts and watermark.ts still define their own copies of the moved types (structurally identical); they were not changed.
3. page_setup and watermark are still NOT wired into apply.ts. No new apply tests.
4. Watermark display is still unverified in Word (TEST_STATUS_v14 caveat 1).

## NOT TESTED
As in TEST_STATUS_v14, plus: wiring of page setup and watermark into apply, partial apply (3c), header/footer choices (5a), new section insertion, captions, cross-references, TOC, tables, images, notes, report, UI, template store, image reduction, batch (not built).
