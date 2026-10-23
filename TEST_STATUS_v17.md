# formatIQ TEST STATUS v17 (chat 17; update on top of TEST_STATUS_v12 to v16; read with HANDOFF_v29)

Method: sandbox rebuilt from scratch (Linux, 1 CPU). Node 22.22.2 and Node 18.20.8. Vitest 2.1.9, TypeScript 5.9.3, Vite 5.4.21.

## TESTED (really run, chat 17)
- src/core/watermark.ts VERSION 2 (17909 bytes, already in Drive, id 1VFty_ZiM-_bH_WLK-dl07t_Wp_b3lG9A; it was saved by another chat at 13:25 UTC with no test record) with the unchanged watermark.test.ts (14311 bytes): 35 of 35 pass on Node 22 AND Node 18; typecheck clean (strict, noUncheckedIndexedAccess, noUnused*); build ok with a temporary entry (bundle 22.95 kB; entry, index.html and dist removed); no "any", no non-null "!", ASCII only (searched).
- Sandbox copies were typed from the Drive text WITHOUT comments (types.ts v2, xml_rewrite.ts, package_merge.ts, watermark.ts) and the test file typed in full. Code is the same as Drive apart from comments. Drive byte size 17909 is the commented file.
- Result: watermark.ts v2 now counts as TESTED. The other chat had already moved the old file to "wrong" as watermark_OLD_v28.ts (id 1as0FPD8oSEr3EHdT2ZsBUbkGZ0vvdvl1). v28 NEXT 1 is DONE.
- Not redone: everything in TEST_STATUS_v12 to v16.

## CAVEATS (open; v12 to v16 caveats still apply)
1. Watermark display is still unverified in Word (TEST_STATUS_v14 caveat 1).
2. Importers of types.ts were not typechecked together (apply.test.ts not loaded); FormatOptions fields pageSetup, watermark, keepSourceHeaders are still optional.
3. page_setup and watermark are still NOT wired into apply.ts.
4. The Drive suite has not been run all together.

## NOT TESTED
As in TEST_STATUS_v16.
