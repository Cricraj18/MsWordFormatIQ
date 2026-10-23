# REVIEW_v1 addendum: sandbox runs (chat 6)

Rebuilt from Drive text: zipio.ts and zipio.test.ts (comments left out) plus a TRIMMED types.ts (only the parts zipio uses). package.json as in Drive, tsconfig strict + noUncheckedIndexedAccess + noUnused*, vitest environment node.

RESULTS (really run)
- npm run typecheck (TypeScript 5.9.3): clean, exit 0.
- npm test on Node 22.22.2 (Vitest 2.1.9, zip.js 2.23.0): 15 of 15 pass.
- npm test on Node 18.20.8 (Vitest 2.1.9): 15 of 15 pass. Node 18 was fetched with: npm i node@18 in /tmp (works in the sandbox).
- Node 18 has CompressionStream but NOT "deflate-raw" (probed). Tests still pass, so zip.js falls back by itself. => REVIEW item T1 is NOT needed. Drop it.

PROBES (zip.js 2.23.0)
- Writing "../evil.xml" is allowed. Reading it fails with "Unsafe filename". => L6 confirmed: the unsafe-path test checks zip.js, not isSafeZipPath. mapError regex is correct for 2.23.0.
- "a//b.xml" is accepted by zip.js and by isSafeZipPath. => L2 confirmed (our check is the only guard).
- A non-zip file fails with "File format is not recognized". => L3: the message is this, not "End of central directory...". If NOT_DOCX is wanted, match /not recognized/i.

NOT TESTED: full types.ts (only trimmed copy), vite build, real browser, 1 GB file, items M1/M2 behaviour (not implemented yet), whether zip.js checks size mismatch (M1 still UNSURE).
