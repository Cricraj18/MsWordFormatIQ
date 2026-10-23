# formatIQ TEST STATUS v18 (chat 19; update on top of TEST_STATUS_v12 to v17; read with HANDOFF_v30)

Method: only the sandbox skeleton was rebuilt (package.json, tsconfig.json, vite.config.ts, npm install; Node 22.22.2, Node 18.20.8 at /tmp/n18). NO source file and NO test was loaded or run in this chat.

## TESTED (really run, chat 19)
Nothing new.

## UNTESTED, ON DRIVE (important)
- src/core/apply.ts version 2 (26122 bytes, id 1XT49S-VXtYRsyphgtsYA_yXGC9BmbKWK, saved 2026-10-17 19:01 UTC by another chat with no handoff and no test record): page_setup and watermark wired in. Never typechecked or run. apply.test.ts in Drive is still the old 22087-byte file (21 tests) written for the old 21510-byte apply.ts (now apply_OLD_v29.ts in "wrong"). Do not count apply.ts as tested until the test run is done.
- Not redone: everything in TEST_STATUS_v12 to v17.

## CAVEATS (open; v12 to v17 caveats still apply)
1. Watermark display is still unverified in Word.
2. apply.ts v2 untested (above); the old apply tests may fail against it (for example new report fields, new imports).
3. FormatOptions fields pageSetup, watermark, keepSourceHeaders are still optional; importers not typechecked together.
4. The Drive suite has not been run all together.

## NOT TESTED
As in TEST_STATUS_v17, plus apply.ts v2 and all its new behaviour (page setup and watermark through apply).
