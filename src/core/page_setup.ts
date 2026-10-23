/**
 * formatIQ page setup (file 5b).
 *
 * Page size, orientation and margins for every section of word/document.xml.
 * A section is a w:sectPr: inside the last paragraph of a section (w:pPr) or
 * at the end of w:body. Section breaks themselves (w:type, columns, borders,
 * page numbering, headers and footers, title page flag...) are NEVER touched:
 * only the direct children w:pgSz and w:pgMar change.
 *
 * Pure token transforms for xml_rewrite.ts, no zip access. Streaming: only
 * one w:sectPr (a few tokens) is kept in memory at a time.
 *
 * Rules (tested):
 * - Sizes are in twips (1/1440 inch). Presets are given as PORTRAIT sizes.
 * - preset "template" uses the template's page size; "custom" uses
 *   customWidthTwips x customHeightTwips (both between 2880 and 31680).
 * - orientation "portrait" or "landscape" forces it. "template": a section
 *   that is landscape in the SOURCE stays landscape (wide tables), every
 *   other section follows the template's orientation. Landscape means the
 *   long side is the width and w:orient="landscape" is written. Portrait
 *   writes no w:orient. w:code (paper code) is removed because the size may
 *   no longer match it.
 * - margins "template" copies the template's first-section margins (only the
 *   values the template has). A MarginSet sets all seven values. Margin values
 *   do not rotate with the page: they are written as they are.
 * - sections "all" or a list of section numbers (0 = first section).
 * - Missing w:pgSz / w:pgMar are created in the correct place of w:sectPr.
 *
 * Version 2 (chat 16): PagePreset, Orientation, MarginSet and PageSetupOptions
 * now live in ../types and are re-exported from here, so old imports still work.
 */
import { FormatIQError } from "../types";
import type { MarginSet, Orientation, PagePreset, PageSetupOptions } from "../types";
import { closeTag, openTag } from "./xml_rewrite";
import type { RawAttr, RawOpen, RawToken, RawTransform } from "./xml_rewrite";
import { createXmlRewriter } from "./xml_rewrite";

export type { MarginSet, Orientation, PagePreset, PageSetupOptions };

/** Portrait sizes of the presets, in twips. */
export const PAGE_SIZES: Readonly<Record<"a4" | "letter" | "a3" | "legal" | "a5" | "tabloid", { width: number; height: number }>> = {
  a4: { width: 11906, height: 16838 },
  letter: { width: 12240, height: 15840 },
  a3: { width: 16838, height: 23811 },
  legal: { width: 12240, height: 20160 },
  a5: { width: 8392, height: 11906 },
  tabloid: { width: 15840, height: 24480 },
};

/** Smallest and largest page side Word accepts (2 inches and 22 inches). */
export const MIN_PAGE_TWIPS = 2880;
export const MAX_PAGE_TWIPS = 31680;

/** The text area must stay at least this wide or high (half an inch). */
export const MIN_TEXT_TWIPS = 720;

/** Default options: everything from the template, all sections. */
export function createDefaultPageSetup(): PageSetupOptions {
  return { preset: "template", orientation: "template", margins: "template", sections: "all" };
}

/** Millimetres to twips (rounded). */
export function mmToTwips(mm: number): number {
  return Math.round((mm / 25.4) * 1440);
}

/** Inches to twips (rounded). */
export function inchesToTwips(inches: number): number {
  return Math.round(inches * 1440);
}

/** What one section says about its page. Missing values are undefined. */
export interface SectionFacts {
  /** 0 for the first section in file order. */
  index: number;
  widthTwips: number | undefined;
  heightTwips: number | undefined;
  /** w:orient="landscape", or width larger than height when orient is missing. */
  landscape: boolean;
  margins: Partial<MarginSet>;
  /** w:type value (nextPage, continuous...) or undefined. */
  breakType: string | undefined;
}

const MARGIN_FIELDS: readonly (keyof MarginSet)[] = ["top", "right", "bottom", "left", "header", "footer", "gutter"];

function localOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function prefixOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? "" : name.slice(0, colon + 1);
}

function attrByLocal(tag: RawOpen, local: string): string | undefined {
  for (const [key, value] of tag.attrs) {
    if (!key.startsWith("xmlns") && localOf(key) === local) return value;
  }
  return undefined;
}

function toInt(value: string | undefined): number | undefined {
  if (value === undefined || !/^-?\d+$/.test(value.trim())) return undefined;
  const n = Number.parseInt(value.trim(), 10);
  return Number.isSafeInteger(n) ? n : undefined;
}

/** Copy of attrs with the attribute (found by local name) set; added with the prefix when new. */
function setAttr(attrs: readonly RawAttr[], local: string, value: string, prefix: string): RawAttr[] {
  let found = false;
  const out = attrs.map((pair): RawAttr => {
    if (!pair[0].startsWith("xmlns") && localOf(pair[0]) === local) {
      found = true;
      return [pair[0], value];
    }
    return pair;
  });
  if (!found) out.push([prefix + local, value]);
  return out;
}

function removeAttr(attrs: readonly RawAttr[], local: string): RawAttr[] {
  return attrs.filter((pair) => pair[0].startsWith("xmlns") || localOf(pair[0]) !== local);
}

/** Direct children of a buffered w:sectPr: name, first and last token index. */
interface ChildRange {
  local: string;
  start: number;
  end: number;
}

function childRanges(tokens: readonly RawToken[]): ChildRange[] {
  const ranges: ChildRange[] = [];
  let depth = 0;
  let current: ChildRange | undefined;
  tokens.forEach((token, i) => {
    if (token.kind === "open") {
      depth += 1;
      if (depth === 2) current = { local: localOf(token.name), start: i, end: i };
    } else if (token.kind === "close") {
      if (depth === 2 && current !== undefined) {
        current.end = i;
        ranges.push(current);
        current = undefined;
      }
      depth -= 1;
    }
  });
  return ranges;
}

/** Read the page facts of one buffered w:sectPr (tokens from its open to its close tag). */
export function factsFromSectPr(index: number, tokens: readonly RawToken[]): SectionFacts {
  const facts: SectionFacts = { index, widthTwips: undefined, heightTwips: undefined, landscape: false, margins: {}, breakType: undefined };
  let orient: string | undefined;
  let depth = 0;
  for (const token of tokens) {
    if (token.kind === "close") {
      depth -= 1;
      continue;
    }
    if (token.kind !== "open") continue;
    depth += 1;
    if (depth !== 2) continue;
    const local = localOf(token.name);
    if (local === "pgSz") {
      facts.widthTwips = toInt(attrByLocal(token, "w"));
      facts.heightTwips = toInt(attrByLocal(token, "h"));
      orient = attrByLocal(token, "orient");
    } else if (local === "pgMar") {
      for (const field of MARGIN_FIELDS) {
        const value = toInt(attrByLocal(token, field));
        if (value !== undefined) facts.margins[field] = value;
      }
    } else if (local === "type") {
      facts.breakType = attrByLocal(token, "val");
    }
  }
  if (orient === "landscape") facts.landscape = true;
  else if (orient === undefined && facts.widthTwips !== undefined && facts.heightTwips !== undefined) {
    facts.landscape = facts.widthTwips > facts.heightTwips;
  }
  return facts;
}

/**
 * Transform that finds every outermost w:sectPr (also inside paragraph
 * properties), hands its tokens to onSection and emits what that returns.
 * A w:sectPr nested inside another one (sectPrChange record) is part of the
 * outer one and is never handled separately.
 */
export function createSectionTransform(onSection: (index: number, tokens: RawToken[]) => RawToken[]): RawTransform {
  let index = -1;
  let buffer: RawToken[] | null = null;
  let depth = 0;
  return (token, emit) => {
    if (buffer !== null) {
      buffer.push(token);
      if (token.kind === "open") depth += 1;
      else if (token.kind === "close") depth -= 1;
      if (token.kind === "close" && depth === 0) {
        const done = buffer;
        buffer = null;
        for (const t of onSection(index, done)) emit(t);
      }
      return;
    }
    if (token.kind === "open" && localOf(token.name) === "sectPr") {
      index += 1;
      buffer = [token];
      depth = 1;
      return;
    }
    emit(token);
  };
}

/** Facts of every section of an XML text (document.xml of a template or a source). */
export function scanSections(xml: string, partName = "word/document.xml"): SectionFacts[] {
  const list: SectionFacts[] = [];
  const rewriter = createXmlRewriter({
    partName,
    onOutput: () => undefined,
    transform: createSectionTransform((index, tokens) => {
      list.push(factsFromSectPr(index, tokens));
      return [];
    }),
  });
  rewriter.write(xml);
  rewriter.end();
  return list;
}

/** What one section gets. margins has only the values to write. */
export interface PageSetupPlan {
  width: number;
  height: number;
  landscape: boolean;
  margins: Partial<MarginSet>;
}

function invalid(message: string, detail?: string): FormatIQError {
  return new FormatIQError("INVALID_OPTIONS", message, detail);
}

/** Problems with a margin set on a page of this size, in plain words. Empty = fine. */
export function checkMargins(margins: Partial<MarginSet>, pageWidth: number, pageHeight: number): string[] {
  const problems: string[] = [];
  for (const field of MARGIN_FIELDS) {
    const value = margins[field];
    if (value === undefined) continue;
    if (!Number.isFinite(value) || !Number.isInteger(value) || value < 0) {
      problems.push(`The ${field} margin must be a whole number of twips, zero or more.`);
    }
  }
  const left = margins.left ?? 0;
  const right = margins.right ?? 0;
  const gutter = margins.gutter ?? 0;
  const top = margins.top ?? 0;
  const bottom = margins.bottom ?? 0;
  if (pageWidth - left - right - gutter < MIN_TEXT_TWIPS) problems.push("The left and right margins leave no room for text.");
  if (pageHeight - top - bottom < MIN_TEXT_TWIPS) problems.push("The top and bottom margins leave no room for text.");
  return problems;
}

/** The portrait (short x long) size for the chosen preset. */
export function portraitSize(options: PageSetupOptions, template: SectionFacts | undefined): { short: number; long: number } {
  let a: number | undefined;
  let b: number | undefined;
  if (options.preset === "custom") {
    a = options.customWidthTwips;
    b = options.customHeightTwips;
    if (a === undefined || b === undefined || !Number.isInteger(a) || !Number.isInteger(b)) {
      throw invalid("Enter a width and a height for the custom page size.");
    }
    if (a < MIN_PAGE_TWIPS || b < MIN_PAGE_TWIPS || a > MAX_PAGE_TWIPS || b > MAX_PAGE_TWIPS) {
      throw invalid("The page must be between 2 and 22 inches on each side.");
    }
  } else if (options.preset === "template") {
    a = template?.widthTwips;
    b = template?.heightTwips;
    if (a === undefined || b === undefined) {
      a = PAGE_SIZES.a4.width;
      b = PAGE_SIZES.a4.height;
    }
  } else {
    a = PAGE_SIZES[options.preset].width;
    b = PAGE_SIZES[options.preset].height;
  }
  return { short: Math.min(a, b), long: Math.max(a, b) };
}

/**
 * What a section gets, or undefined when the section is not selected.
 * template = facts of the template's first section (may be undefined for a
 * template without one: A4 portrait is used then). source = the section's own facts.
 * Throws FormatIQError INVALID_OPTIONS for bad custom sizes or margins.
 */
export function planSection(options: PageSetupOptions, template: SectionFacts | undefined, source: SectionFacts): PageSetupPlan | undefined {
  if (options.sections !== "all" && !options.sections.includes(source.index)) return undefined;
  const size = portraitSize(options, template);
  let landscape: boolean;
  if (options.orientation === "portrait") landscape = false;
  else if (options.orientation === "landscape") landscape = true;
  else landscape = source.landscape || (template?.landscape ?? false);
  const width = landscape ? size.long : size.short;
  const height = landscape ? size.short : size.long;
  const margins: Partial<MarginSet> = options.margins === "template" ? { ...(template?.margins ?? {}) } : { ...options.margins };
  const problems = checkMargins(margins, width, height);
  if (problems.length > 0) {
    throw invalid(problems[0] ?? "The margins are not valid.", `section ${source.index + 1}`);
  }
  return { width, height, landscape, margins };
}

/** Counters filled by createPageSetupTransform. */
export interface PageSetupStats {
  sections: number;
  changed: number;
}

/** Children that come before w:pgSz in w:sectPr (schema order). */
const BEFORE_PGSZ: ReadonlySet<string> = new Set(["headerReference", "footerReference", "footnotePr", "endnotePr", "type"]);

/** Rewrite w:pgSz and w:pgMar of one buffered w:sectPr. Returns new tokens. */
export function applyPlanToSectPr(tokens: readonly RawToken[], plan: PageSetupPlan): RawToken[] {
  const first = tokens[0];
  if (first === undefined || first.kind !== "open") return [...tokens];
  const prefix = prefixOf(first.name);
  const ranges = childRanges(tokens);
  const pgSz = ranges.find((r) => r.local === "pgSz");
  const pgMar = ranges.find((r) => r.local === "pgMar");
  const hasMargins = MARGIN_FIELDS.some((f) => plan.margins[f] !== undefined);

  const makeSize = (old: RawOpen | undefined): RawToken[] => {
    let attrs: RawAttr[] = old === undefined ? [] : [...old.attrs];
    attrs = setAttr(attrs, "w", String(plan.width), prefix);
    attrs = setAttr(attrs, "h", String(plan.height), prefix);
    attrs = plan.landscape ? setAttr(attrs, "orient", "landscape", prefix) : removeAttr(attrs, "orient");
    attrs = removeAttr(attrs, "code");
    return [openTag(old?.name ?? prefix + "pgSz", attrs), closeTag(old?.name ?? prefix + "pgSz")];
  };
  const makeMargins = (old: RawOpen | undefined): RawToken[] => {
    let attrs: RawAttr[] = old === undefined ? [] : [...old.attrs];
    for (const field of MARGIN_FIELDS) {
      const value = plan.margins[field];
      if (value !== undefined) attrs = setAttr(attrs, field, String(value), prefix);
    }
    return [openTag(old?.name ?? prefix + "pgMar", attrs), closeTag(old?.name ?? prefix + "pgMar")];
  };

  const oldSz = pgSz === undefined ? undefined : tokens[pgSz.start];
  const oldMar = pgMar === undefined ? undefined : tokens[pgMar.start];
  const sz = makeSize(oldSz !== undefined && oldSz.kind === "open" ? oldSz : undefined);
  const mar = hasMargins || pgMar !== undefined ? makeMargins(oldMar !== undefined && oldMar.kind === "open" ? oldMar : undefined) : [];
  const keepOldMar = !hasMargins && pgMar !== undefined;

  // Where new elements go when w:pgSz is missing: after the last "before" child.
  let insertAt = 1;
  for (const r of ranges) {
    if (BEFORE_PGSZ.has(r.local)) insertAt = r.end + 1;
  }

  const out: RawToken[] = [];
  let i = 0;
  while (i < tokens.length) {
    const token = tokens[i];
    if (token === undefined) break;
    if (pgSz !== undefined && i === pgSz.start) {
      out.push(...sz);
      if (pgMar === undefined) out.push(...mar);
      i = pgSz.end + 1;
      continue;
    }
    if (pgMar !== undefined && i === pgMar.start) {
      if (keepOldMar) out.push(...tokens.slice(pgMar.start, pgMar.end + 1));
      else out.push(...mar);
      i = pgMar.end + 1;
      continue;
    }
    if (pgSz === undefined && i === insertAt) {
      out.push(...sz);
      if (pgMar === undefined) out.push(...mar);
    }
    out.push(token);
    i += 1;
  }
  return out;
}

/**
 * Transform that applies the page setup to the chosen sections of document.xml.
 * template = facts of the template's first section (from scanSections on the
 * template's document.xml, first entry), or undefined.
 */
export function createPageSetupTransform(
  options: PageSetupOptions,
  template: SectionFacts | undefined,
): { transform: RawTransform; stats: PageSetupStats } {
  const stats: PageSetupStats = { sections: 0, changed: 0 };
  const transform = createSectionTransform((index, tokens) => {
    stats.sections += 1;
    const plan = planSection(options, template, factsFromSectPr(index, tokens));
    if (plan === undefined) return tokens;
    stats.changed += 1;
    return applyPlanToSectPr(tokens, plan);
  });
  return { transform, stats };
}
