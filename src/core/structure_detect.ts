/**
 * formatIQ structure detection (file 2a, part 3).
 *
 * Reads the RAW document in two streaming passes (nothing big is kept):
 *  pass 1: count font sizes to find the usual body size (small map only);
 *  pass 2: classify every paragraph with the rules in structure_rules.ts and
 *          hand each result to a callback.
 * The raw file is only read, never changed.
 *
 * Two ways to use it:
 *  - detectStructureStream(): callback per paragraph. Use this for big files
 *    (the formatter should, so it never holds a million results).
 *  - detectStructure(): collects all results into a DetectionResult (types.ts).
 *    Fine for small and medium files, and for tests.
 * Both work on a Blob (a File) or on an already opened docx.
 */
import { LOW_CONFIDENCE_THRESHOLD } from "../types";
import type { DetectedParagraph, DetectionResult, ParagraphInfo, ProgressCallback } from "../types";
import { DEFAULT_ZIP_LIMITS, openDocx } from "./zipio";
import type { JobControl, OpenedDocx, ZipLimits } from "./zipio";
import { readParagraphs, readStyleHints } from "./paragraphs";
import { FALLBACK_BODY_SIZE, classifyParagraph, pickBodySize } from "./structure_rules";
import type { DetectContext } from "./structure_rules";

/** Optional settings. */
export interface DetectOptions {
  limits?: ZipLimits;
  control?: JobControl;
  onProgress?: ProgressCallback;
}

/** Numbers about a finished detection. */
export interface DetectionSummary {
  paragraphCount: number;
  lowConfidenceCount: number;
  /** Body font size that was used, in half points. */
  bodySizeHalfPoints: number;
}

/** Detect on an opened docx. The caller closes it. */
export async function detectStructureInDocx(
  docx: OpenedDocx,
  onDetected: (item: DetectedParagraph) => void,
  options: DetectOptions = {},
): Promise<DetectionSummary> {
  const limits = options.limits ?? DEFAULT_ZIP_LIMITS;
  const report = (fraction: number, message: string): void => {
    options.onProgress?.({ phase: "detecting", fraction, message });
  };

  report(0, "Reading styles");
  const hints = await readStyleHints(docx, limits, options.control);

  report(0.1, "Finding the body text size");
  const histogram = new Map<number, number>();
  await readParagraphs(
    docx,
    (p) => {
      if (p.inTable || p.text.trim() === "") return;
      const hint = p.styleId === null ? undefined : hints.styles.get(p.styleId);
      if (hint !== undefined && (hint.outlineLevel !== null || /^(heading|title|subtitle)/i.test(hint.name))) return;
      if (p.outlineLevel !== null) return;
      const size = p.sizeHalfPoints ?? hint?.sizeHalfPoints ?? hints.defaultSizeHalfPoints ?? FALLBACK_BODY_SIZE;
      histogram.set(size, (histogram.get(size) ?? 0) + p.text.length);
    },
    limits,
    options.control,
  );
  const bodySize = pickBodySize(histogram) ?? hints.defaultSizeHalfPoints ?? FALLBACK_BODY_SIZE;

  report(0.5, "Detecting structure");
  const ctx: DetectContext = {
    styles: hints.styles,
    bodySizeHalfPoints: bodySize,
    defaultSizeHalfPoints: hints.defaultSizeHalfPoints,
    nonEmptyBefore: 0,
    headingSeen: false,
  };
  let count = 0;
  let low = 0;
  await readParagraphs(
    docx,
    (p) => {
      const result = classifyParagraph(p, ctx);
      const info: ParagraphInfo = {
        index: p.index, text: p.text, styleId: p.styleId, inTable: p.inTable,
        hasImage: p.hasImage, isListItem: p.isListItem,
      };
      if (p.text.trim() !== "") ctx.nonEmptyBefore += 1;
      if (result.role === "heading" && result.confidence >= LOW_CONFIDENCE_THRESHOLD) ctx.headingSeen = true;
      if (result.confidence < LOW_CONFIDENCE_THRESHOLD) low += 1;
      count += 1;
      onDetected({ paragraph: info, role: result.role, level: result.level, confidence: result.confidence, manualRole: null });
    },
    limits,
    options.control,
  );
  report(1, "Structure detected");
  return { paragraphCount: count, lowConfidenceCount: low, bodySizeHalfPoints: bodySize };
}

/** Stream results for a file (Blob). Opens and closes the docx itself. */
export async function detectStructureStream(
  source: Blob,
  fileName: string,
  onDetected: (item: DetectedParagraph) => void,
  options: DetectOptions = {},
): Promise<DetectionSummary> {
  const docx = await openDocx(source, fileName, options.limits ?? DEFAULT_ZIP_LIMITS, options.control);
  try {
    return await detectStructureInDocx(docx, onDetected, options);
  } finally {
    await docx.close().catch(() => undefined);
  }
}

/** Collect all results into one DetectionResult (types.ts). */
export async function detectStructure(
  source: Blob,
  fileName: string,
  options: DetectOptions = {},
): Promise<DetectionResult> {
  const paragraphs: DetectedParagraph[] = [];
  const summary = await detectStructureStream(source, fileName, (item) => { paragraphs.push(item); }, options);
  return { paragraphs, lowConfidenceCount: summary.lowConfidenceCount };
}
