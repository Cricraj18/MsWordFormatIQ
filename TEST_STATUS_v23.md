# TEST_STATUS v23 (2026-10-18)
Typecheck (tsc --noEmit): clean. Node 22.22 vitest: 19 files, 417 tests passed. Node 18.20.8 vitest: 19 files, 417 tests passed. Build (vite): ok.
First real run of apply.ts (apply.test.ts 21) and apply_extras.test.ts (16): all passed, no code fix needed.
Changed: types.ts (pageSetup, watermark, keepSourceHeaders required), apply.test.ts and types.test.ts (one test each). Re-run after the change: same 417 passed on both Node versions.
Not tested: real browser, real Word rendering, 1 GB file.
