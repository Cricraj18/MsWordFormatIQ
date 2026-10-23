# formatIQ HANDOFF v5

FOLDER RULES
- Project folder: FORMAT_IQ, ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA. Subfolder src, ID 108HasgSfuacTIwGOV-ugUoqXmMp6nkCl. Subfolder src/core, ID 1gD5Q2WEWwISUiOoFhZ-kblS0pMPjlIxf. Code files go in their real folders (src/, src/core/, src/modules/). Create subfolders if missing.
- Drive tools cannot overwrite file content. Each update is saved as a NEW file: HANDOFF_v6.md, progress_v5.md, and so on. Always use the highest number.
- Outdated, wrong or corrupt files: MOVE them into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda). Never delete anything. After saving a new HANDOFF_vN, move the previous HANDOFF_v(N-1) and older progress files to "wrong".
- If moving or renaming fails (needs user approval), tell the user in one line and continue.
- The user said: it is OK to build several files in one chat, up to the chat limit, but ALWAYS stop at a TEST GATE.

You are a senior TypeScript developer helping a beginner build "formatIQ", a browser-only Word (.docx) formatting tool. Everything runs in the browser. No server, no uploads, no AI, no paid APIs.

RULES
- Build ONLY the NEXT FILE(S) below, in order, and stop at the TEST GATE.
- Follow SPEC exactly. If unclear, ask before guessing.
- Complete files only, under about 300 lines. Add doc comments, error handling, Vitest tests.
- Never modify the user's original file. Output is a new file.
- Native Word fields only (SEQ, REF, TOC, PAGE, NUMPAGES, STYLEREF).
- No invented APIs: write "UNSURE" and say how to check.
- No "any" and no "!" (non-null assertion) in TypeScript.
- I am a beginner. Explain steps simply, in short messages (the user reads on a phone).
- The user has NO computer at home. They test only at the office on working days. So you cannot assume anything has been run unless the DONE list says "tested".

SPEC (short)
- Template .docx is the base package. Raw content is poured into it, mapped to template styles.
- Files up to 1 GB: lazy zip reading (zip.js), streaming XML, media copied untouched, streamed output, Web Worker, progress and Cancel.
- Safety: reject non-.docx/.docm, limit unzipped and document.xml size, reject unsafe zip paths, no XML entity expansion, time limit.
- Preserve unknown parts unchanged (comments, footnotes, SmartArt, equations, custom XML, tracked changes).
- Structure detection: rules only, confidence score, manual tagging for low confidence.
- Partial apply: whole template or selected parts.
- Cross-references: SEQ and REF fields with bookmarks, TOC, lists of figures/tables. Set update-fields-on-open and warn that fields may look empty until updated.
- Sections: tag-only landscape with Next-Page breaks, swapped page size, headers/footers and page numbers continuous.
- Image reduction: optional separate module, OFF by default. Original / Light (~2400 px, Q90, ~1 MB) / Balanced (~1600 px, Q90, ~500 KB). Pixels only, never upscale, keep displayed size and crop, min 150 DPI, skip line art and transparent images, one image at a time in a worker, progress and Cancel.
- Templates stored in browser (IndexedDB) or as exported files.
- Non-goals: PDF tools, OCR, AutoCAD, login, database, mobile apps, AI.

FILE LIST (in order)
0a package.json, tsconfig, Vite, Vitest setup [DONE, not yet run]
0b src/types.ts [DONE, not yet run]
0c src/core/zipio.ts (lazy read) [DONE, not yet run]
0d src/core/zipio.ts (streamed write) [DONE, not yet run]
--- TEST GATE: STOP. Do not build 1a until the user reports that 0a to 0d were run at the office (npm install, npm install @zip.js/zip.js, typecheck, test, dev) and what happened. ---
1a template_extract.ts 1b its tests
2a structure_detect.ts 2b its tests
3a-3c formatter (styles, apply, partial)
4a-4b crossref (SEQ, REF)
5a-5c toc, headers/footers, sections
6 manual_mark.ts
7 web UI
8 template_store.ts
9 image_reduce.ts
10 deploy guide

DONE
- 0a (written, NOT yet run). Files in FORMAT_IQ: package.json, tsconfig.json, vite.config.ts, index.html, and src/ with version.ts, main.ts, version.test.ts.
  Decisions: ESM ("type": "module"). TypeScript strict plus noUncheckedIndexedAccess, noUnusedLocals, noUnusedParameters. tsconfig lib has DOM and WebWorker. vite.config.ts uses defineConfig from "vitest/config". Tests sit next to code as src/**/*.test.ts, environment "node". Scripts: npm run dev / build / typecheck / test. Imports in src have no file extensions.
  Package versions (typescript ^5.5, vite ^5.4, vitest ^2.0) are UNSURE.
  NOTE: package.json does NOT yet list @zip.js/zip.js. User must run `npm install @zip.js/zip.js` at the office (Drive cannot overwrite package.json).
- 0b (written, NOT yet run). File: src/types.ts (about 230 lines, no imports). No separate test file.
  Exports (names to reuse):
  - Progress: ProgressPhase ("reading" | "detecting" | "formatting" | "images" | "writing" | "done"), ProgressInfo { phase, fraction 0..1, message }, ProgressCallback = (info: ProgressInfo) => void.
  - Errors: FormatIQErrorCode (NOT_DOCX, FILE_TOO_LARGE, XML_TOO_LARGE, UNSAFE_ZIP_PATH, XML_ENTITY_EXPANSION, TIMEOUT, CANCELLED, ZIP_READ_FAILED, ZIP_WRITE_FAILED, TEMPLATE_INVALID, INVALID_OPTIONS, UNKNOWN). class FormatIQError extends Error { code; detail: string | undefined } constructor(code, message, detail?). Functions isFormatIQError(value), toFormatIQError(value).
  - Template: StyleKind, TemplateStyle { id, name, kind }, PageSetup { widthTwips, heightTwips, landscape }, TemplateInfo { id, name, createdAt, styles, page, hasHeaderFooter, hasToc }.
  - Detection: StructureRole ("title" | "heading" | "body" | "list-item" | "figure-caption" | "table-caption" | "table" | "quote" | "code" | "unknown"), ParagraphInfo { index, text, styleId, inTable, hasImage, isListItem }, DetectedParagraph { paragraph, role, level, confidence, manualRole }, LOW_CONFIDENCE_THRESHOLD = 0.6, DetectionResult { paragraphs, lowConfidenceCount }.
  - Options: ImageReductionLevel ("original" | "light" | "balanced"), FormatPart ("styles" | "headersFooters" | "toc" | "crossReferences" | "sections"), FormatOptions { scope, parts, includeToc, includeListOfFigures, includeListOfTables, updateFieldsOnOpen, imageReduction }, DEFAULT_FORMAT_OPTIONS (image reduction "original").
- 0c + 0d (written, NOT yet run). Files: src/core/zipio.ts (about 330 lines, a little over the 300 target) and src/core/zipio.test.ts. Imports from "../types" and from "@zip.js/zip.js" (BlobReader, BlobWriter, TextReader, ZipReader, ZipWriter, configure; types Entry, FileEntry).
  Exports (names to reuse):
  - Limits: ZipLimits { maxEntries, maxTotalBytes, maxXmlBytes, timeoutMs }, DEFAULT_ZIP_LIMITS (50,000 entries, 4 GiB total, 256 MiB per XML part, 10 minutes).
  - Checks: hasAllowedExtension(fileName), isSafeZipPath(path), setWorkersEnabled(enabled).
  - Job control: JobControl { signal, timedOut(), dispose() }, createJobControl(timeoutMs, userSignal?). Cancel = abort the user signal; errors become CANCELLED or TIMEOUT.
  - Read: openDocx(source: Blob, fileName, limits?) returns OpenedDocx { entries: Map<string, FileEntry>, totalBytes, close() }. streamEntry(entry, onChunk, maxBytes, control?). readEntryText(entry, maxBytes, control?).
  - Write: ZipOutput { addText(path, text, control?), copyEntry(entry, control?), finish() }. createZipOutput(target: WritableStream<Uint8Array>, onProgress?) for streaming to disk. createBlobZipOutput(onProgress?) returns { output, getBlob() } for in-memory output.
  Decisions: media and other already-compressed parts are copied by decompress then store (level 0), so the BYTES are identical but it is not a raw copy. Write paths are checked too. [Content_Types].xml should be written first by the caller (convention, not enforced).
  UNSURE (check at office with typecheck and test): (1) entry.getData accepts a WritableStream; (2) ZipWriter accepts a WritableStream; (3) ZipWriter<Blob> is assignable to ZipWriter<unknown>; (4) Entry is a union with a "directory" flag so FileEntry narrows; (5) zip.js option "passThrough" for true raw copy (not used); (6) zip.js works in Node 18+ for the tests; (7) default limits fit a 1 GB file.

NEXT FILE: TEST GATE. Stop. Test at the office first. After the user reports results of 0a to 0d, fix any problems (put replaced files in "wrong"), then build 1a src/core/template_extract.ts (read template .docx: styles.xml, page setup, headers/footers, TOC; fill TemplateInfo from src/types.ts using openDocx and readEntryText from src/core/zipio.ts).

WHAT TO DO IN EVERY CHAT
1. Read this file (highest HANDOFF_vN) and progress_vN. Build the NEXT FILE(S), stop at a TEST GATE.
2. Give: file name and folder, install commands, how to test, expected output, known limits, anything UNSURE. Keep it short.
3. Save the file to the right Drive folder. If you cannot save, say "I can't save to Drive" and give the full file.
4. Save a NEW HANDOFF_v(N+1).md (this text with DONE and NEXT FILE updated, and any changed names or types recorded) and a NEW progress_v(N+1).md. Move the old HANDOFF and old progress files to "wrong".
5. FINISH WITH THE RESUME PROMPT below in one code block, exactly as written, so the user can paste it into a new chat. If the TEST GATE is next, say clearly: "Stop. Test at the office first," and still give the prompt for use after testing.

RESUME PROMPT (give this exact text at the end of every chat)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully. Follow it exactly: build ONLY the NEXT FILE, do not redo anything in DONE, save the file into the right folder, save a new HANDOFF_v(N+1).md and progress_v(N+1).md, move the old HANDOFF to the folder "wrong", and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
