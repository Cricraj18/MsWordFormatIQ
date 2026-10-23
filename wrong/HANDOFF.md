# formatIQ HANDOFF

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
0a package.json, tsconfig, Vite, Vitest setup
0b src/types.ts
0c src/core/zipio.ts (lazy read)
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

DONE: none
NEXT FILE: 0a

AFTER THE CODE
1. Give file name, folder, install commands, test steps, expected output, known limits.
2. Save to Drive only if you have access. If not, say "I can't save to Drive" and give me the full file plus updated progress.md and HANDOFF.md text.
3. List anything marked UNSURE.
4. Rewrite HANDOFF.md with this file added to DONE and the next file as NEXT FILE.
