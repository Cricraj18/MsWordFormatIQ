# formatIQ code review v1 (types.ts, zipio.ts, zipio.test.ts)

Method: read the 3 files from Drive. NOT re-run in sandbox in this chat (Node 22 only; Node 18 not available). Line numbers for zipio.ts are counted by hand from the Drive text, so "about" +/- 2. Test file: located by test name.
Checked: no "any" found. No "!" non-null assertions found (all "!" are "!==" or "!expr"). Handoff v6 says typecheck was clean and 15 zipio tests passed on Node 22 / zip.js 2.23.0 / Vitest 2.1.9.
No type errors found by reading. No wrong zip.js calls found: getData(WritableStream), getData(BlobWriter), ZipWriter(WritableStream), add(name, Reader, {level,lastModDate,signal,onprogress}), getEntries(), close() all match the v2 API as I know it. UNSURE only where marked.

## MEDIUM

M1. zipio.ts about L319-325 (copyEntry): no byte cap. The zip index can lie (zip bomb), and getData(new BlobWriter()) then holds the real size in memory.
  Fix (partial, after L319): add
    if (part.size !== entry.uncompressedSize) throw new FormatIQError("ZIP_READ_FAILED", "Part size does not match the zip index.", entry.filename);
  This catches lies but memory is already used. A true cap needs a capped WritableStream, which risks the deadlock noted in HANDOFF_v6. UNSURE whether zip.js already rejects size mismatch (its checkSignature option may do CRC checks; default value UNSURE).

M2. zipio.ts about L163-208 (openDocx): ZipLimits.timeoutMs is never used anywhere, and getEntries() cannot be cancelled (no JobControl param).
  Fix: add 4th param control?: JobControl; call reader.getEntries(signalOptions(control)) (UNSURE: whether getEntries accepts "signal" in 2.23; check with npm run typecheck); in the catch use mapError(err, control, "ZIP_READ_FAILED").

M3. DEFAULT_FORMAT_OPTIONS in types.ts (last lines): "parts: []" is one shared array. Any code that does DEFAULT_FORMAT_OPTIONS.parts.push(...) changes the default for everyone.
  Fix: replace the const with a function: export function createDefaultFormatOptions(): FormatOptions { return { scope: "whole", parts: [], includeToc: true, includeListOfFigures: false, includeListOfTables: false, updateFieldsOnOpen: true, imageReduction: "original" }; } (keep the old const only if nothing uses it yet).

## LOW

L1. zipio.ts about L95-98 (createJobControl): setTimeout delay above 2147483647 ms fires at once (browsers and Node). Fix: setTimeout(..., Math.min(timeoutMs, 2_147_483_647)).
L2. zipio.ts about L74 (isSafeZipPath): "a//b" passes (empty segment). Fix: path.replace(/\/$/, "").split("/").every((p) => p !== "" && p !== ".." && p !== "."). Test "word/media/" still passes.
L3. zipio.ts about L204-206: a file that is not a zip at all (for example renamed .txt) gives ZIP_READ_FAILED, not NOT_DOCX. Decide which you want. UNSURE: exact zip.js message ("End of central directory not found" as I remember).
L4. zipio.ts L8-9, L213, L344: UNSURE comments are out of date. getData(WritableStream) and ZipWriter(WritableStream) were verified by tests. Fix: reword to "Tested with zip.js 2.23.0."
L5. zipio.ts about L276 (isAlreadyCompressed): embedded xlsx/pptx/pdf/heic get recompressed at level 6. Fix: add xlsx|pptx|pdf|heic to the regex.
L6. zipio.ts about L129 and test "rejects an unsafe path inside the zip": the "../evil.xml" case is probably rejected by zip.js itself ("Unsafe filename" text match), so isSafeZipPath inside openDocx (L180) may never run for it. UNSURE. The test then tests zip.js wording, not our check. Fix: keep the test, but do not rely on it for isSafeZipPath (it has its own unit tests).
L7. Cancel/timeout in the middle of writing leaves the ZipWriter and the target stream open. UNSURE if zip.js offers a clean abort. Decide at UI time (file 7).

## TESTS (Node 18+ / Vitest)

T1. Node 18 UNSURE: zip.js uses native CompressionStream when it exists. Node 18 has the global, but I could not verify "deflate-raw" there. Safe fix in beforeAll: import { configure } from "@zip.js/zip.js"; configure({ useWebWorkers: false, useCompressionStream: false });. (Or add useCompressionStream:false inside setWorkersEnabled in zipio.ts, only for tests.) Run on a real Node 18 to confirm.
T2. Test "writes into a WritableStream": chunks are copied at the END (.map(c => new Uint8Array(c))). If zip.js ever reuses a buffer, the data is wrong. Fix: write(chunk) { chunks.push(chunk.slice()); } and keep the .map(...) (it also avoids the TS 5.7+ Uint8Array<ArrayBufferLike> vs BlobPart error).
T3. Test "copies a large binary part byte for byte": the name is wrong. A 32-bit hash is not byte for byte, and the LCG multiplication passes 2^53 (loses precision, still repeatable).
  Fix 1: x = (Math.imul(x, 1103515245) + 12345) & 0x7fffffff;
  Fix 2: replace checksum with a direct compare: let offset = 0; await streamEntry(first, (chunk) => { for (let i = 0; i < chunk.length; i += 1) { if (chunk[i] !== big[offset + i]) throw new Error("Byte " + (offset + i) + " differs"); } offset += chunk.length; }, big.length + 10); expect(offset).toBe(big.length);
T4. Same test: await import("@zip.js/zip.js") is not needed. Fix: add ZipReader and BlobReader to the import at the top.
T5. Test "reports CANCELLED when the job is already aborted" depends on zip.js checking the signal before it finishes. It passed on 2.23.0. package.json says ^2.23.0, so older versions are excluded. OK.

## Files state in Drive (checked)
- Complete (not cut off): src/types.ts, src/core/zipio.ts, src/core/zipio.test.ts, HANDOFF_v6.md.
- Missing: progress_v5.md / progress_v6.md (last chat stopped before saving).
- Still in the project root but outdated (should be in "wrong"): HANDOFF_v5.md (id 11avZjBBzHigkhB7P7aE9KlciHGuDA_va), progress_v4.md (id 14tEt2DSQtr-gPd3VMuxcZAaq3qgYPAwj). The Drive tools here have no move/rename. Move them by hand.
