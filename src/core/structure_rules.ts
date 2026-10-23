/**
 * formatIQ structure rules (file 2a, part 2). Rules only, no AI.
 *
 * Small pure functions (easy to test) and classifyParagraph(), which combines
 * them into a role + level + confidence for one paragraph.
 * Confidence below LOW_CONFIDENCE_THRESHOLD (0.6) means "ask the user".
 *
 * Order of the rules in classifyParagraph (first match wins):
 *  1 inside a table -> table
 *  2 no text -> body (a picture alone is body, 0.8)
 *  3 Word style: Heading N, Title, Subtitle, Quote, Code -> that role (0.95)
 *  4 outline level (paragraph or style) -> heading (0.9)
 *  5 caption: "Figure 3", "Table 7", or style Caption
 *  6 text patterns: "Chapter 4", "Annexure A", "1 Intro", "2.1 Scope"
 *  7 list numbering -> list-item (0.9)
 *  8 looks: big bold short line (title or heading, 0.5 to 0.7), bold short line
 *  9 otherwise body (0.85)
 * Guesses that come only from looks stay near or below 0.6 on purpose.
 */
import type { StructureRole } from "../types";
import type { ParagraphFacts, StyleHint } from "./paragraphs";

/** Role, heading level and confidence for one paragraph. */
export interface RuleResult {
  role: StructureRole;
  level: number | null;
  confidence: number;
}

/** What classifyParagraph needs besides the paragraph itself. */
export interface DetectContext {
  /** Paragraph styles by id. */
  styles: ReadonlyMap<string, StyleHint>;
  /** Most common body font size in half points (22 = 11 pt). */
  bodySizeHalfPoints: number;
  /** Document default font size in half points, or null. */
  defaultSizeHalfPoints: number | null;
  /** How many paragraphs WITH text came before this one. */
  nonEmptyBefore: number;
  /** True once a heading with confidence of 0.6 or more was found. */
  headingSeen: boolean;
}

/** Default body size when nothing is known: 11 pt. */
export const FALLBACK_BODY_SIZE = 22;

const clampLevel = (n: number): number => Math.min(9, Math.max(1, n));

const WORD_NUMBERS = "one|two|three|four|five|six|seven|eight|nine|ten";

function wordCount(text: string): number {
  const t = text.trim();
  return t === "" ? 0 : t.split(/\s+/).length;
}

/** "1 Introduction", "2.1 Scope", "3.2.1. Details". Level = number of parts. */
export function matchNumberedHeading(text: string): { level: number; confidence: number } | null {
  const m = /^(\d{1,3}(?:\.\d{1,3}){0,5})\.?\s+(\S.*)$/u.exec(text.trim());
  if (m === null) return null;
  const parts = (m[1] ?? "").split(".").length;
  const rest = m[2] ?? "";
  if (rest.length > 100 || wordCount(rest) > 12) return null;
  if (/[.,;]$/.test(rest) || /^\p{Ll}/u.test(rest)) return null;
  if (parts >= 2) return { level: clampLevel(parts), confidence: 0.85 };
  return { level: 1, confidence: wordCount(rest) <= 6 ? 0.7 : 0.55 };
}

/** "Chapter 4", "Part II", "Section 3", "Annexure A", "Appendix B". */
export function matchLabeledHeading(text: string): { level: number; confidence: number } | null {
  const t = text.trim();
  const m = new RegExp(`^(chapter|part|unit|section)\\s+(\\d+|[ivxlcdm]+|${WORD_NUMBERS})\\b(.*)$`, "iu").exec(t);
  const a = /^(annexure|annex|appendix|schedule)\s+([A-Za-z]|\d+|[ivxIVX]+)\b(.*)$/iu.exec(t);
  const hit = m ?? a;
  if (hit === null) return null;
  const word = (hit[1] ?? "").toLowerCase();
  const rest = hit[3] ?? "";
  const level = word === "section" ? 2 : 1;
  if (t.length > 100 || wordCount(rest) > 10) return null;
  if (/[.,;]$/.test(rest.trim()) || /^\s+\p{Ll}/u.test(rest)) return { level, confidence: 0.5 };
  return { level, confidence: 0.9 };
}

/** "Figure 3: ...", "Fig. 2.1", "Table 7 - ...". Longer sentences are not captions. */
export function matchCaption(text: string): { role: "figure-caption" | "table-caption"; confidence: number } | null {
  const t = text.trim();
  const m = /^(figure|fig\.?|table)\s*(\d+(?:[.\-]\d+)*[a-z]?|[IVX]+)(.*)$/iu.exec(t);
  if (m === null || t.length > 250) return null;
  const word = (m[1] ?? "").toLowerCase();
  const rest = m[3] ?? "";
  const role = word === "table" ? "table-caption" : "figure-caption";
  if (rest.trim() === "" || /^\s*[:.\-\u2013\u2014]/u.test(rest)) return { role, confidence: 0.92 };
  if (/^\s+\p{Ll}/u.test(rest)) return { role, confidence: 0.4 };
  return { role, confidence: t.length <= 150 ? 0.8 : 0.55 };
}

/** Role from a Word style name or id ("heading 2", "Title", "Quote" ...). */
export function roleFromStyle(styleId: string | null, hint: StyleHint | undefined): { role: StructureRole; level: number | null } | null {
  const name = (hint?.name ?? styleId ?? "").toLowerCase().trim();
  if (name === "") return null;
  const h = /^heading\s*([1-9])$/.exec(name);
  if (h !== null) return { role: "heading", level: Number(h[1]) };
  if (name === "title" || name === "subtitle") return { role: "title", level: null };
  if (/^(quote|intense quote|block text)$/.test(name)) return { role: "quote", level: null };
  if (/^(source code|code|html preformatted|macro text)$/.test(name)) return { role: "code", level: null };
  return null;
}

/** True for a short line without sentence punctuation at the end. */
export function isShortLine(text: string, maxChars: number): boolean {
  const t = text.trim();
  return t.length > 0 && t.length <= maxChars && !/[.!?,;:]$/.test(t);
}

/** Heading level guess from the size ratio to body text. */
export function levelFromSizeRatio(ratio: number): number {
  return ratio >= 1.6 ? 1 : ratio >= 1.3 ? 2 : 3;
}

/** Find the most common size (counted by characters). Returns null if empty. */
export function pickBodySize(histogram: ReadonlyMap<number, number>): number | null {
  let best: number | null = null;
  let bestWeight = 0;
  for (const [size, weight] of histogram) {
    if (weight > bestWeight || (weight === bestWeight && best !== null && size < best)) {
      best = size;
      bestWeight = weight;
    }
  }
  return best;
}

/** Decide the role of one paragraph. See the order in the file comment. */
export function classifyParagraph(p: ParagraphFacts, ctx: DetectContext): RuleResult {
  if (p.inTable) return { role: "table", level: null, confidence: 0.95 };
  const text = p.text.trim();
  if (text === "") return { role: "body", level: null, confidence: p.hasImage ? 0.8 : 1 };

  const hint = p.styleId === null ? undefined : ctx.styles.get(p.styleId);
  const fromStyle = roleFromStyle(p.styleId, hint);
  if (fromStyle !== null) return { role: fromStyle.role, level: fromStyle.level, confidence: 0.95 };

  const outline = p.outlineLevel ?? hint?.outlineLevel ?? null;
  if (outline !== null) return { role: "heading", level: clampLevel(outline), confidence: 0.9 };

  const styleName = (hint?.name ?? p.styleId ?? "").toLowerCase();
  const caption = matchCaption(text);
  if (caption !== null) return { role: caption.role, level: null, confidence: caption.confidence };
  if (styleName === "caption") {
    return { role: "figure-caption", level: null, confidence: 0.55 };
  }

  const bold = p.bold ?? hint?.bold ?? false;
  const size = p.sizeHalfPoints ?? hint?.sizeHalfPoints ?? ctx.defaultSizeHalfPoints ?? ctx.bodySizeHalfPoints;
  const ratio = size / ctx.bodySizeHalfPoints;

  const labeled = matchLabeledHeading(text);
  if (labeled !== null) {
    const boost = bold || ratio >= 1.15 ? 0.05 : 0;
    return { role: "heading", level: labeled.level, confidence: Math.min(0.95, labeled.confidence + boost) };
  }
  if (!p.isListItem) {
    const numbered = matchNumberedHeading(text);
    if (numbered !== null) {
      const boost = bold || ratio >= 1.15 ? 0.1 : 0;
      return { role: "heading", level: numbered.level, confidence: Math.min(0.95, numbered.confidence + boost) };
    }
  }

  if (p.isListItem) return { role: "list-item", level: null, confidence: 0.9 };

  if (isShortLine(text, 80)) {
    if (ratio >= 1.8 && ctx.nonEmptyBefore < 5 && !ctx.headingSeen) return { role: "title", level: null, confidence: 0.7 };
    if (bold && ratio >= 1.15) return { role: "heading", level: levelFromSizeRatio(ratio), confidence: 0.65 };
    if (!bold && ratio >= 1.5) return { role: "heading", level: levelFromSizeRatio(ratio), confidence: 0.5 };
  }
  if (bold && ratio < 1.15 && isShortLine(text, 60)) return { role: "unknown", level: null, confidence: 0.35 };

  return { role: "body", level: null, confidence: 0.85 };
}
