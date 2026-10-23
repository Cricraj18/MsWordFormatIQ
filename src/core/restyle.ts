/**
 * formatIQ restyler (file 3b-1).
 *
 * A token transform for xml_rewrite.ts. Fed with the tokens of
 * word/document.xml, it gives each top-level paragraph the paragraph style
 * chosen by the caller and, if asked, removes direct formatting that would
 * hide the new style.
 *
 * PARAGRAPH NUMBERING (must match paragraphs.ts, tested): paragraph
 * numbers 0, 1, 2... are given to every w:p that is NOT inside another w:p,
 * in file order. Paragraphs in tables and content controls count. Paragraphs
 * inside text boxes (nested in a w:p) are not counted and are not changed.
 *
 * What changes, for a paragraph whose styleFor(index) is a style id:
 * - w:pStyle becomes the first child of w:pPr (schema order). An old
 *   w:pStyle is dropped. A missing w:pPr is created.
 * - directFormatting "clear" (default for the product) also removes, from
 *   the paragraph itself: jc, spacing, ind, contextualSpacing, mirrorIndents,
 *   shd, pBdr; and from every run and the paragraph mark: rFonts, sz, szCs,
 *   color, kern, spacing, w, position. KEPT: numPr, keepNext, keepLines,
 *   pageBreakBefore, tabs, outlineLvl, sectPr (section breaks!), bold, italic,
 *   underline, strike, sub/superscript, highlight, run shading, rStyle,
 *   language, fields, hyperlinks, breaks, images.
 * - Records of tracked formatting changes (w:rPrChange, w:pPrChange) are
 *   never touched.
 * When styleFor gives null the paragraph is copied exactly as it is.
 *
 * The new elements use the same name prefix as the paragraph ("w:" in Word
 * files). Streaming: only a few counters are kept.
 */
import { closeTag, openTag } from "./xml_rewrite";
import type { RawOpen, RawToken, RawTransform } from "./xml_rewrite";

/** "clear" removes direct formatting that hides the style; "keep" only sets the style. */
export type DirectFormatting = "clear" | "keep";

/** What the restyler needs from the caller. */
export interface RestyleOptions {
  /** Style id for the paragraph with this number, or null to leave it alone. */
  styleFor: (index: number) => string | null;
  directFormatting: DirectFormatting;
}

/** Counters, filled while the tokens go through. */
export interface RestyleStats {
  /** Top-level paragraphs seen. */
  paragraphs: number;
  /** Paragraphs that got a style. */
  restyled: number;
  /** Paragraph property elements removed. */
  clearedParagraphProps: number;
  /** Run property elements removed. */
  clearedRunProps: number;
}

/** The transform to give to createXmlRewriter, and its counters. */
export interface Restyler {
  transform: RawTransform;
  stats: RestyleStats;
}

/** Direct paragraph formatting removed in "clear" mode (local names). */
export const PARAGRAPH_PROPS_CLEARED: ReadonlySet<string> = new Set([
  "jc",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "shd",
  "pBdr",
]);

/** Direct run formatting removed in "clear" mode (local names). */
export const RUN_PROPS_CLEARED: ReadonlySet<string> = new Set([
  "rFonts",
  "sz",
  "szCs",
  "color",
  "kern",
  "spacing",
  "w",
  "position",
]);

function localOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function prefixOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? "" : name.slice(0, colon + 1);
}

/** Make the restyler. Make a new one for every document part. */
export function createRestyler(options: RestyleOptions): Restyler {
  const stats: RestyleStats = { paragraphs: 0, restyled: 0, clearedParagraphProps: 0, clearedRunProps: 0 };
  const clearMode = options.directFormatting === "clear";
  const stack: string[] = [];
  let paragraphIndex = -1;
  let pDepth = 0;
  let changeDepth = 0;
  let skipDepth = 0;
  /** Style for the current top-level paragraph, null = leave alone. */
  let style: string | null = null;
  let prefix = "";
  /** True from the start of an active paragraph until its first child. */
  let needPPr = false;
  /** True once the paragraph's own w:pPr was handled or created. */
  let pPrDone = false;
  /** Stack length while inside the paragraph's own w:pPr, else -1. */
  let ownPPrLen = -1;

  const topLocal = (): string => {
    const top = stack[stack.length - 1];
    return top === undefined ? "" : localOf(top);
  };

  const writeStyle = (emit: (token: RawToken) => void): void => {
    if (style === null) return;
    emit(openTag(prefix + "pStyle", [[prefix + "val", style]]));
    emit(closeTag(prefix + "pStyle"));
    stats.restyled += 1;
  };

  const insertPPr = (emit: (token: RawToken) => void): void => {
    emit(openTag(prefix + "pPr"));
    writeStyle(emit);
    emit(closeTag(prefix + "pPr"));
    needPPr = false;
    pPrDone = true;
  };

  const skip = (token: RawOpen): void => {
    stack.push(token.name);
    skipDepth = 1;
  };

  const onOpen = (token: RawOpen, emit: (token: RawToken) => void): void => {
    const local = localOf(token.name);
    const parent = topLocal();
    const directChildOfParagraph = pDepth === 1 && parent === "p";

    if (needPPr && directChildOfParagraph) {
      if (local !== "pPr") insertPPr(emit);
      else needPPr = false;
    }

    if (local === "p") {
      pDepth += 1;
      if (pDepth === 1) {
        paragraphIndex += 1;
        stats.paragraphs += 1;
        style = options.styleFor(paragraphIndex);
        prefix = prefixOf(token.name);
        needPPr = style !== null;
        pPrDone = false;
        ownPPrLen = -1;
      }
      emit(token);
      stack.push(token.name);
      return;
    }

    if (local === "rPrChange" || local === "pPrChange") {
      changeDepth += 1;
      emit(token);
      stack.push(token.name);
      return;
    }

    const active = style !== null && pDepth === 1 && changeDepth === 0;

    if (active && local === "pPr" && directChildOfParagraph && !pPrDone) {
      pPrDone = true;
      emit(token);
      stack.push(token.name);
      ownPPrLen = stack.length;
      writeStyle(emit);
      return;
    }

    if (active && parent === "pPr" && stack.length === ownPPrLen) {
      if (local === "pStyle") {
        skip(token);
        return;
      }
      if (clearMode && PARAGRAPH_PROPS_CLEARED.has(local)) {
        stats.clearedParagraphProps += 1;
        skip(token);
        return;
      }
    }

    if (active && clearMode && parent === "rPr" && RUN_PROPS_CLEARED.has(local)) {
      stats.clearedRunProps += 1;
      skip(token);
      return;
    }

    emit(token);
    stack.push(token.name);
  };

  const transform: RawTransform = (token, emit) => {
    if (skipDepth > 0) {
      if (token.kind === "open") {
        stack.push(token.name);
        skipDepth += 1;
      } else if (token.kind === "close") {
        stack.pop();
        skipDepth -= 1;
      }
      return;
    }
    if (token.kind === "open") {
      onOpen(token, emit);
      return;
    }
    if (token.kind === "close") {
      const local = localOf(token.name);
      if (local === "p") {
        if (pDepth === 1 && needPPr) insertPPr(emit);
        if (pDepth === 1) {
          style = null;
          ownPPrLen = -1;
        }
        pDepth -= 1;
      } else if (local === "rPrChange" || local === "pPrChange") {
        changeDepth -= 1;
      } else if (local === "pPr" && stack.length === ownPPrLen) {
        ownPPrLen = -1;
      }
      stack.pop();
      emit(token);
      return;
    }
    if (needPPr && pDepth === 1 && topLocal() === "p") {
      const blank = token.kind === "raw" || token.text.trim() === "";
      if (!blank) insertPPr(emit);
    }
    emit(token);
  };

  return { transform, stats };
}
