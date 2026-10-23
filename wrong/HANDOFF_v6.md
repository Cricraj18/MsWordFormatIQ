# formatIQ HANDOFF v6

FOLDER RULES
- Project folder: FORMAT_IQ, ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA. Subfolder src, ID 108HasgSfuacTIwGOV-ugUoqXmMp6nkCl. Subfolder src/core, ID 1gD5Q2WEWwISUiOoFhZ-kblS0pMPjlIxf. Code files go in their real folders (src/, src/core/, src/modules/). Create subfolders if missing.
- Drive tools cannot overwrite file content. Each update is saved as a NEW file: HANDOFF_v7.md, progress_v6.md, and so on. Always use the highest number. For code files, if a file must be replaced: move the old one to "wrong" (rename it with _OLD_vN) and create the new one with the real name.
- Outdated, wrong or corrupt files: MOVE them into the folder named "wrong" (ID 1306nkF0ShJox257WewKZ2s6QK51fzgda). Never delete anything. After saving a new HANDOFF_vN, move the previous HANDOFF_v(N-1) and older progress files to "wrong".
- If moving or renaming fails (needs user approval), tell the user in one line and continue.
- The user said: it is OK to build several files in one chat, up to the chat limit, but ALWAYS stop at a TEST GATE.

HOW CLAUDE CAN TEST (IMPORTANT, learned in chat 5)
- Claude HAS a Linux sandbox (bash_tool) with Node 22 and npm access. So Claude CAN and MUST run npm install, npm run typecheck, npm test and npm run build for every file before saving. Never say "I cannot run it".
- To rebuild the project in the sandbox: read code files from Drive with download_file_content (it returns base64; replace the text "\u003d" with "=" before decoding). Do NOT use read_file_content for code: it adds backslashes before brackets and stars. Then write them under /home/claude/formatiq, run npm install, and test.
- The sandbox cannot test a real browser (Web Workers, file picker, dev server page). Those need the user at the office.
- Verified sandbox versions: Node 22.22, npm 10.9, TypeScript 5.9.3, Vite 5.4.21, Vitest 2.1.9, @zip.js/zip.js 2.23.0.

You are a senior TypeScript developer helping a beginner build "formatIQ", a browser-only Word (.docx) formatting tool. Everything runs in the browser. No server, no uploads, no AI, no paid APIs.

RULES
- Build ONLY the NEXT FILE(S) below, in order, and stop at the TEST GATE.
- Follow SPEC exactly. If unclear, ask before guessing.
- Complete files only, under about 300 lines. Add doc comments, error handling, Vitest tests.
- Never modify the user's original file. Output is a new file.
- Native Word fields only (SEQ, REF, TOC, PAGE, NUMPAGES, STYLEREF).
- No invented APIs: write "UNSURE" and say how to check. Better: test it in the sandbox.
- No "any" and no "!" (non-null assertion) in TypeScript.
- I am a beginner. Explain steps simply, in short messages (the user reads on a phone).
- Say "tested" only for what was really run. Sandbox runs count as "tested in sandbox". Browser behaviour is not tested yet.

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
0a package.json, tsconfig, Vite, Vitest setup [DONE, tested in sandbox]
0b src/types.ts [DONE, tested in sandbox]
0c src/core/zipio.ts (lazy read) [DONE, tested in sandbox]
0d src/core/zipio.ts (write) [DONE, tested in sandbox]
--- TEST GATE: PASSED in sandbox (typecheck clean, 17 tests pass, build ok). Optional at the office: npm install, npm run dev (page should show "formatIQ v0.0.1"). Do not wait for it. ---
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
- 0a (tested in sandbox). Files in FORMAT_IQ: package.json (NEW copy now includes dependency "@zip.js/zip.js": "^2.23.0"; old copy is in "wrong" as package_OLD_v1.json), tsconfig.json, vite.config.ts, index.html, and src/ with version.ts, main.ts, version.test.ts.
  Decisions: ESM ("type": "module"). TypeScript strict plus noUncheckedIndexedAccess, noUnusedLocals, noUnusedParameters. tsconfig target ES2022, lib has DOM and WebWorker. vite.config.ts uses defineConfig from "vitest/config". Tests sit next to code as src/**/*.test.ts, environment "node". Scripts: npm run dev / build / typecheck / test. Imports in src have no file extensions.
  npm run build works in the sandbox.
- 0b (tested in sandbox by typecheck; no own test file). File: src/types.ts (232 lines, no imports).
  Exports (names to reuse):
  - Progress: ProgressPhase ("reading" | "detecting" | "formatting" | "images" | "writing" | "done"), ProgressInfo { phase, fraction 0..1, message }, ProgressCallback = (info: ProgressInfo) => void.
  - Errors: FormatIQErrorCode (NOT_DOCX, FILE_TOO_LARGE, XML_TOO_LARGE, UNSAFE_ZIP_PATH, XML_ENTITY_EXPANSION, TIMEOUT, CANCELLED, ZIP_READ_FAILED, ZIP_WRITE_FAILED, TEMPLATE_INVALID, INVALID_OPTIONS, UNKNOWN). class FormatIQError extends Error { code; detail: string | undefined } constructor(code, message, detail?). Functions isFormatIQError(value), toFormatIQError(value).
  - Template: StyleKind, TemplateStyle { id, name, kind }, PageSetup { widthTwips, heightTwips, landscape }, TemplateInfo { id, name, createdAt, styles, page, hasHeaderFooter, hasToc }.
  - Detection: StructureRole ("title" | "heading" | "body" | "list-item" | "figure-caption" | "table-caption" | "table" | "quote" | "code" | "unknown"), ParagraphInfo { index, text, styleId, inTable, hasImage, isListItem }, DetectedParagraph { paragraph, role, level, confidence, manualRole }, LOW_CONFIDENCE_THRESHOLD = 0.6, DetectionResult { paragraphs, lowConfidenceCount }.
  - Options: ImageReductionLevel ("original" | "light" | "balanced"), FormatPart ("styles" | "headersFooters" | "toc" | "crossReferences" | "sections"), FormatOptions { scope, parts, includeToc, includeListOfFigures, includeListOfTables, updateFieldsOnOpen, imageReduction }, DEFAULT_FORMAT_OPTIONS (image reduction "original").
- 0c + 0d (tested in sandbox: 15 zipio tests pass). Files: src/core/zipio.ts (364 lines, over the 300 target; split later if it grows) and src/core/zipio.test.ts. Imports from "../types" and "@zip.js/zip.js".
  Exports (names to reuse):
  - Limits: ZipLimits { maxEntries, maxTotalBytes, maxXmlBytes, timeoutMs }, DEFAULT_ZIP_LIMITS (50,000 entries, 4 GiB total, 256 MiB per XML part, 10 minutes).
  - Checks: hasAllowedExtension(fileName), isSafeZipPath(path), setWorkersEnabled(enabled) (tests call setWorkersEnabled(false) in beforeAll).
  - Job control: JobControl { signal, timedOut(), dispose() }, createJobControl(timeoutMs, userSignal?). Cancel = abort the user signal; errors become CANCELLED or TIMEOUT.
  - Read: openDocx(source: Blob, fileName, limits?) returns OpenedDocx { entries: Map<string, FileEntry>, totalBytes, close() }. streamEntry(entry, onChunk, maxBytes, control?). readEntryText(entry, maxBytes, control?).
  - Write: ZipOutput { addText(path, text, control?), copyEntry(entry, control?), finish() }. createZipOutput(target: WritableStream<Uint8Array>, onProgress?) for streaming to a stream. createBlobZipOutput(onProgress?) returns { output, getBlob() } for in-memory output.
  Facts learned by real testing (zip.js 2.23.0):
  - Reading with getData(WritableStream) works. Writing into a WritableStream works. ZipWriter<Blob> is accepted where ZipWriter<unknown> is expected. FileEntry narrows with entry.directory.
  - DEADLOCK: running entry.getData(pipe) and zipWriter.add(name, pipe.readable) at the same time hangs (TransformStream, and a manual bridge too). So copyEntry copies ONE part at a time: getData into a Blob, then add from that Blob. Memory = one part, not the whole file. A single part bigger than available memory is a known limit.
  - Copied parts keep identical bytes (tested with 3 MB random data) but are decompressed and written again; images are stored at level 0. Not a raw copy. The zip.js option "passThrough" was not tried.
  - zip.js itself rejects unsafe names while reading (message "Unsafe filename"). zipio maps it to UNSAFE_ZIP_PATH. Wording could change in other versions.
  - Write paths are also checked (UNSAFE_ZIP_PATH). Write [Content_Types].xml first (convention, not enforced).
  UNSURE (not testable in the sandbox): zip.js Web Workers in a real browser; streaming to disk with showSaveFilePicker; default limits for a real 1 GB file; real speed.

NEXT FILE: 1a (src/core/template_extract.ts: read a template .docx with openDocx and readEntryText from src/core/zipio.ts; fill a TemplateInfo from src/types.ts: styles from word/styles.xml, page setup from the first w:sectPr in word/document.xml, hasHeaderFooter from header/footer parts, hasToc from a TOC field). Then 1b (its tests, src/core/template_extract.test.ts). Build both in this order, run them in the sandbox, then continue to 2a only if the chat has room. Choose an XML approach that works in browser AND Node tests and is safe (no entity expansion): check what is available (DOMParser is browser only; UNSURE for Node tests). Ask the user only if truly unclear.

WHAT TO DO IN EVERY CHAT
1. Read this file (highest HANDOFF_vN) and progress_vN. Build the NEXT FILE(S), stop at a TEST GATE.
2. Rebuild the project in the sandbox and RUN typecheck and tests before saving anything (see HOW CLAUDE CAN TEST).
3. Give: file name and folder, install commands, how to test, expected output, known limits, anything UNSURE. Keep it short.
4. Save the file to the right Drive folder. If you cannot save, say "I can't save to Drive" and give the full file.
5. Save a NEW HANDOFF_v(N+1).md (this text with DONE and NEXT FILE updated, and any changed names or types recorded) and a NEW progress_v(N+1).md. Move the old HANDOFF and old progress files to "wrong".
6. FINISH WITH THE RESUME PROMPT below in one code block, exactly as written, so the user can paste it into a new chat. If a TEST GATE is next, say clearly: "Stop. Test at the office first," and still give the prompt for use after testing.

RESUME PROMPT (give this exact text at the end of every chat)
My Google Drive folder "FORMAT_IQ" is here: https://drive.google.com/drive/folders/17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA (folder ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA). Open the highest-numbered HANDOFF_vN.md in that folder and read it fully. Follow it exactly: build ONLY the NEXT FILE, do not redo anything in DONE, save the file into the right folder, save a new HANDOFF_v(N+1).md and progress_v(N+1).md, move the old HANDOFF to the folder "wrong", and end by giving me this same resume prompt in a code block. If you cannot open or write to the folder, say so plainly and give me the full file to save by hand. Do not pretend to save.
