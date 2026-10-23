# formatIQ TEST STATUS v16 (chat 16; update on top of TEST_STATUS_v12 to v15; read with HANDOFF_v28)

Method: sandbox rebuilt from scratch (Linux, 1 CPU). Node 22.22.2 and Node 18.20.8. Vitest 2.1.9, TypeScript 5.9.3, Vite 5.4.21.

## TESTED (really run, chat 16)
- src/core/page_setup.ts v2 (16360 bytes, saved in Drive) with the unchanged page_setup.test.ts (16895 bytes): 39 of 39 pass on Node 22 AND Node 18; typecheck clean (strict, noUncheckedIndexedAccess, noUnused*); build ok with a temporary entry (bundle 28.34 kB; entry and index.html removed); no "any", no non-null "!", ASCII only (searched). Only change from the old file: the four shared types are imported from ../types and re-exported. Drive copy has the same byte size as the sandbox copy.
- The sandbox had only types.ts (comment-free copy of Drive v2), xml_rewrite.ts (comment-free) and the page_setup files. Other suites were NOT rerun.
- Not redone: everything in TEST_STATUS_v12 to v15.

## CAVEATS (open; v12 to v15 caveats still apply)
1. Importers of types.ts were not typechecked together (apply.test.ts not loaded).
2. watermark.ts still defines its own copies of WatermarkMode and WatermarkOptions.
3. page_setup and watermark are still NOT wired into apply.ts.

## NOT TESTED
As in TEST_STATUS_v15.
