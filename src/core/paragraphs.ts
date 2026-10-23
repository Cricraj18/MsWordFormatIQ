/**
 * formatIQ streaming paragraph reader (file 2a, part 1).
 *
 * Turns word/document.xml into one ParagraphFacts object per paragraph, one at
 * a time, through a callback. No array of all paragraphs is ever built here, so
 * a document with a million paragraphs stays small in memory. Text is cut at
 * MAX_PARAGRAPH_TEXT characters (enough for "1.2 Scope", "Figure 3", ...).
 *
 * Also reads the paragraph styles from word/styles.xml into a small map
 * (name, outline level, bold, font size, with "based on" followed).
 *
 * Simple rules, decided and tested:
 * - Only top-level paragraphs count. A paragraph nested inside another one
 *   (a text box) is skipped, and its text is not added to the outer paragraph.
 * - Text comes from w:t only (not deleted text w:delText, not field codes).
 *   w:tab gives a tab character, w:br and w:cr give one space.
 * - inTable is true for every paragraph inside any w:tbl (also nested tables).
 * - isListItem is true when the paragraph itself has w:numPr with numId not 0.
 *   (A list set only by a style is NOT seen. numbering.xml is not read.)
 * - bold: true when every run that has text is explicitly bold; false when one
 *   run is explicitly not bold, or bold is only on some runs; null when no run
 *   says anything (then the style decides). sizeHalfPoints: the biggest w:sz of
 *   a run that has text, or null.
 */
import type { ParagraphInfo } from "../types";
import { DEFAULT_ZIP_LIMITS } from "./zipio";
import type { JobControl, OpenedDocx, ZipLimits } from "./zipio";
import { parseEntryXml } from "./xml";
import type { XmlHandlers } from "./xml";

/** Most characters of text kept per paragraph. */
export const MAX_PARAGRAPH_TEXT = 300;

/** ParagraphInfo plus the formatting facts that detection needs. */
export interface ParagraphFacts extends ParagraphInfo {
  /** Outline level 1 to 9 set on the paragraph itself, or null. */
  outlineLevel: number | null;
  /** See the file comment. */
  bold: boolean | null;
  /** Biggest explicit run font size in half points (22 = 11 pt), or null. */
  sizeHalfPoints: number | null;
}

/** w:b / w:i style on-off value: missing or 1/true/on is true. */
function onOff(value: string | undefined): boolean {
  if (value === undefined) return true;
  const v = value.toLowerCase();
  return !(v === "0" || v === "false" || v === "off");
}

function intOrNull(value: string | undefined): number | null {
  if (value === undefined || !/^\d+$/.test(value)) return null;
  const n = Number(value);
  return Number.isSafeInteger(n) ? n : null;
}

interface Builder {
  text: string;
  styleId: string | null;
  inTable: boolean;
  hasImage: boolean;
  isListItem: boolean;
  outlineLevel: number | null;
  hasTrue: boolean;
  hasFalse: boolean;
  hasNull: boolean;
  size: number | null;
}

/** Make the XML handlers that call emit() once per top-level paragraph. */
export function createParagraphHandlers(emit: (facts: ParagraphFacts) => void): XmlHandlers {
  let index = 0;
  let tblDepth = 0;
  let pDepth = 0;
  let cur: Builder | null = null;
  let inPPr = false;
  let inRun = false;
  let inRPr = false;
  let inT = false;
  let runHasText = false;
  let runBold: boolean | null = null;
  let runSize: number | null = null;

  const add = (piece: string): void => {
    if (cur === null || cur.text.length >= MAX_PARAGRAPH_TEXT) return;
    cur.text = (cur.text + piece).slice(0, MAX_PARAGRAPH_TEXT);
  };

  return {
    open(tag) {
      if (tag.name === "p") {
        pDepth += 1;
        if (pDepth === 1) {
          cur = { text: "", styleId: null, inTable: tblDepth > 0, hasImage: false, isListItem: false, outlineLevel: null, hasTrue: false, hasFalse: false, hasNull: false, size: null };
          inPPr = false; inRun = false; inRPr = false; inT = false;
        }
        return;
      }
      if (tag.name === "tbl") {
        tblDepth += 1;
        return;
      }
      if (pDepth !== 1 || cur === null) return;
      switch (tag.name) {
        case "pPr": inPPr = true; break;
        case "pStyle": if (inPPr) cur.styleId = tag.attrs["val"] ?? null; break;
        case "numId": if (inPPr && tag.attrs["val"] !== undefined && tag.attrs["val"] !== "0") cur.isListItem = true; break;
        case "outlineLvl": {
          const n = intOrNull(tag.attrs["val"]);
          if (inPPr && n !== null && n <= 8) cur.outlineLevel = n + 1;
          break;
        }
        case "r": inRun = true; runHasText = false; runBold = null; runSize = null; break;
        case "rPr": if (!inPPr && inRun) inRPr = true; break;
        case "b": if (inRPr) runBold = onOff(tag.attrs["val"]); break;
        case "sz": if (inRPr) runSize = intOrNull(tag.attrs["val"]); break;
        case "t": if (inRun) inT = true; break;
        case "tab": if (inRun && !inPPr) add("\t"); break;
        case "br":
        case "cr": if (inRun && !inPPr) add(" "); break;
        case "drawing":
        case "pict": cur.hasImage = true; break;
        default: break;
      }
    },
    text(text) {
      if (pDepth !== 1 || !inT) return;
      if (text.trim() !== "") runHasText = true;
      add(text);
    },
    close(name) {
      if (name === "tbl") { tblDepth -= 1; return; }
      if (name === "p") {
        if (pDepth === 1 && cur !== null) {
          const bold = cur.hasFalse || (cur.hasTrue && cur.hasNull) ? false : cur.hasTrue ? true : null;
          emit({
            index, text: cur.text, styleId: cur.styleId, inTable: cur.inTable, hasImage: cur.hasImage,
            isListItem: cur.isListItem, outlineLevel: cur.outlineLevel, bold, sizeHalfPoints: cur.size,
          });
          index += 1;
          cur = null;
        }
        pDepth -= 1;
        return;
      }
      if (pDepth !== 1 || cur === null) return;
      switch (name) {
        case "t": inT = false; break;
        case "rPr": inRPr = false; break;
        case "pPr": inPPr = false; break;
        case "r":
          if (inRun && runHasText) {
            if (runBold === true) cur.hasTrue = true;
            else if (runBold === false) cur.hasFalse = true;
            else cur.hasNull = true;
            if (runSize !== null && (cur.size === null || runSize > cur.size)) cur.size = runSize;
          }
          inRun = false;
          break;
        default: break;
      }
    },
  };
}

/** What detection needs to know about one paragraph style. */
export interface StyleHint {
  /** Name shown in Word, for example "heading 1" (id if no name). */
  name: string;
  outlineLevel: number | null;
  bold: boolean | null;
  sizeHalfPoints: number | null;
}

/** Paragraph styles by id, plus the document default font size. */
export interface StyleHints {
  styles: ReadonlyMap<string, StyleHint>;
  defaultSizeHalfPoints: number | null;
}

interface RawStyle {
  id: string;
  name: string;
  basedOn: string | null;
  outlineLevel: number | null;
  bold: boolean | null;
  size: number | null;
}

/** Make the XML handlers that fill raw style data from styles.xml. */
export function createStyleHandlers(raw: Map<string, RawStyle>, defaults: { size: number | null }): XmlHandlers {
  let cur: RawStyle | null = null;
  let inPPr = false;
  let inRPr = false;
  let inDocDefaults = false;
  return {
    open(tag) {
      switch (tag.name) {
        case "docDefaults": inDocDefaults = true; break;
        case "style":
          cur = tag.attrs["type"] === "paragraph" && (tag.attrs["styleId"] ?? "") !== ""
            ? { id: tag.attrs["styleId"] ?? "", name: "", basedOn: null, outlineLevel: null, bold: null, size: null }
            : null;
          break;
        case "name": if (cur !== null && cur.name === "") cur.name = tag.attrs["val"] ?? ""; break;
        case "basedOn": if (cur !== null) cur.basedOn = tag.attrs["val"] ?? null; break;
        case "pPr": inPPr = true; break;
        case "rPr": inRPr = true; break;
        case "outlineLvl": {
          const n = intOrNull(tag.attrs["val"]);
          if (cur !== null && inPPr && n !== null && n <= 8) cur.outlineLevel = n + 1;
          break;
        }
        case "b": if (cur !== null && inRPr) cur.bold = onOff(tag.attrs["val"]); break;
        case "sz": {
          const n = intOrNull(tag.attrs["val"]);
          if (inRPr && cur !== null) cur.size = n;
          else if (inRPr && inDocDefaults && n !== null) defaults.size = n;
          break;
        }
        default: break;
      }
    },
    close(name) {
      switch (name) {
        case "docDefaults": inDocDefaults = false; break;
        case "pPr": inPPr = false; break;
        case "rPr": inRPr = false; break;
        case "style":
          if (cur !== null && !raw.has(cur.id)) raw.set(cur.id, cur);
          cur = null;
          break;
        default: break;
      }
    },
  };
}

/** Follow "based on" (at most 20 steps, safe against loops). */
export function resolveStyleHints(raw: ReadonlyMap<string, RawStyle>, defaultSize: number | null): StyleHints {
  const styles = new Map<string, StyleHint>();
  for (const style of raw.values()) {
    let outline = style.outlineLevel;
    let bold = style.bold;
    let size = style.size;
    const seen = new Set<string>([style.id]);
    let next = style.basedOn;
    for (let step = 0; step < 20 && next !== null && !seen.has(next); step += 1) {
      seen.add(next);
      const parent = raw.get(next);
      if (parent === undefined) break;
      if (outline === null) outline = parent.outlineLevel;
      if (bold === null) bold = parent.bold;
      if (size === null) size = parent.size;
      next = parent.basedOn;
    }
    styles.set(style.id, { name: style.name === "" ? style.id : style.name, outlineLevel: outline, bold, sizeHalfPoints: size });
  }
  return { styles, defaultSizeHalfPoints: defaultSize };
}

/** Read paragraph styles from word/styles.xml. Empty result if there is none. */
export async function readStyleHints(docx: OpenedDocx, limits: ZipLimits = DEFAULT_ZIP_LIMITS, control?: JobControl): Promise<StyleHints> {
  const entry = docx.entries.get("word/styles.xml");
  if (entry === undefined) return { styles: new Map(), defaultSizeHalfPoints: null };
  const raw = new Map<string, RawStyle>();
  const defaults: { size: number | null } = { size: null };
  await parseEntryXml(entry, createStyleHandlers(raw, defaults), limits.maxXmlBytes, control);
  return resolveStyleHints(raw, defaults.size);
}

/** Stream every top-level paragraph of word/document.xml to onParagraph. */
export async function readParagraphs(
  docx: OpenedDocx,
  onParagraph: (facts: ParagraphFacts) => void,
  limits: ZipLimits = DEFAULT_ZIP_LIMITS,
  control?: JobControl,
): Promise<void> {
  const entry = docx.entries.get("word/document.xml");
  if (entry === undefined) return;
  await parseEntryXml(entry, createParagraphHandlers(onParagraph), limits.maxXmlBytes, control);
}
