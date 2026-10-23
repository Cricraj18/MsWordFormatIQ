# formatIQ HANDOFF v2

(VERSION RULE: Drive tools cannot overwrite a file's content. Each update is saved as a NEW file: HANDOFF_v2.md, HANDOFF_v3.md, ... Always use the highest number. Ignore HANDOFF.md, it is outdated. Same for progress_vN.md.)

You are a senior TypeScript developer helping a beginner build "formatIQ", a browser-only Word (.docx) formatting tool. Everything runs in the browser. No server, no uploads, no AI, no paid APIs.

RULES
- Build ONLY the NEXT FILE below. One file per chat.
- Follow SPEC exactly. If unclear, ask before guessing.
- Complete files only, under about 300 lines. Add doc comments, error handling, Vitest tests.
- Never modify the user's original file. Output is a new file.
- Native Word fields only (SEQ, REF, TOC, PAGE, NUMPAGES, STYLEREF).
- No invented APIs: write "UNSURE" and say how to check.
- No "any" and no "!" in TypeScript.
- I am a beginner. Explain steps simply.

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
0a package.json, tsconfig, Vite, Vitest setup  [DONE]
0b src/types.ts
0c src/core/zipio.ts (lazy read, uses @zip.js/zip.js)
0d src/core/zipio.ts (streamed write)
1a template_extract.ts   1b its tests
2a structure_detect.ts   2b its tests
3a-3c formatter (styles, apply, partial)
4a-4b crossref (SEQ, REF)
5a-5c toc, headers/footers, sections
6 manual_mark.ts
7 web UI
8 template_store.ts
9 image_reduce.ts
10 deploy guide

DONE
- 0a (written, NOT yet run by the user). Files in the FORMAT_IQ Drive folder: package.json, tsconfig.json, vite.config.ts, index.html, and src/ folder with version.ts, main.ts, version.test.ts.
  Decisions: ESM project ("type": "module"). TypeScript strict plus noUncheckedIndexedAccess, noUnusedLocals, noUnusedParameters. tsconfig lib includes DOM and WebWorker. vite.config.ts imports defineConfig from "vitest/config". Tests live next to code as src/**/*.test.ts, environment "node". Scripts: npm run dev / build / typecheck / test.
  Ignore "_old_vite.config.ts (ignore)" in Drive, it was a mistake.
  No runtime dependencies yet. zip.js is added in 0c.
  Package versions (typescript ^5.5, vite ^5.4, vitest ^2.0) are UNSURE: user should run npm install and report any warning.

NEXT FILE: 0b (src/types.ts: TemplateInfo, ParagraphInfo, DetectionResult, FormatOptions, ProgressCallback, plus any small shared error type)
Files in src/ must be imported without extensions (Bundler resolution).

AFTER THE CODE
1. Give file name, folder, install commands, test steps, expected output, known limits.
2. Save to the Drive folder FORMAT_IQ (ID 17Zj2i3Es6hQW8t5s3qhIuTkbSVFgYRhA, src folder ID 108HasgSfuacTIwGOV-ugUoqXmMp6nkCl) only if you have access. If not, say "I can't save to Drive" and give me the full file.
3. List anything marked UNSURE.
4. Save a NEW HANDOFF_v3.md (do not overwrite) with this file added to DONE and the next file as NEXT FILE. Also save progress_v3.md.
