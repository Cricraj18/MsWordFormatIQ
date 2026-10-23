/**
 * formatIQ style mapping (file 3a-2).
 *
 * Decides which paragraph style of the TEMPLATE each detected role gets.
 * Pure functions, no zip and no XML. Used by apply.ts (3b).
 *
 * How styles are found: by the style NAME shown in Word ("heading 1",
 * "Title", "Caption", "Normal"...), ignoring case, spaces and dashes. If a
 * template has no matching name, the style id is tried ("Heading1").
 * Localized Word names (for example German "berschrift 1") are only found
 * through the id pattern, because Word keeps English ids in most files.
 *
 * Choices (documented, tested):
 * - body uses "Normal" first, then "Body Text".
 * - figure caption: "Figure Caption", then "Caption". Table caption:
 *   "Table Caption", then "Caption".
 * - heading level N uses "Heading N". If the template stops at level 3 and
 *   the document has level 5, level 3 is used (deepest style that exists).
 * - "table" and "unknown" roles are NOT restyled (null).
 */
import type { DetectedParagraph, StructureRole, TemplateStyle } from "../types";

/** Biggest heading level Word offers. */
export const MAX_HEADING_LEVEL = 9;

/** Style ids per role. null = the template has no matching style. */
export interface StyleMap {
  /** Index 0 is level 1. Always MAX_HEADING_LEVEL long. */
  headings: (string | null)[];
  title: string | null;
  body: string | null;
  listItem: string | null;
  figureCaption: string | null;
  tableCaption: string | null;
  quote: string | null;
  code: string | null;
}

/** What a paragraph really is: the manual tag wins over the guess. */
export interface EffectiveRole {
  role: StructureRole;
  /** Heading level 1..9, or null when role is not "heading". */
  level: number | null;
}

/** Lower case, only letters and digits: "Heading 1" and "heading-1" match. */
export function normalizeStyleName(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findStyle(styles: readonly TemplateStyle[], names: readonly string[], idPattern?: RegExp): string | null {
  const paragraphStyles = styles.filter((s) => s.kind === "paragraph");
  for (const wanted of names) {
    const key = normalizeStyleName(wanted);
    const hit = paragraphStyles.find((s) => normalizeStyleName(s.name) === key);
    if (hit !== undefined) return hit.id;
  }
  if (idPattern !== undefined) {
    const hit = paragraphStyles.find((s) => idPattern.test(s.id));
    if (hit !== undefined) return hit.id;
  }
  return null;
}

/** Look through the template's styles and fill a StyleMap. */
export function buildStyleMap(styles: readonly TemplateStyle[]): StyleMap {
  const headings: (string | null)[] = [];
  for (let level = 1; level <= MAX_HEADING_LEVEL; level += 1) {
    headings.push(findStyle(styles, [`heading ${level}`], new RegExp(`^heading${level}$`, "i")));
  }
  const caption = findStyle(styles, ["caption"], /^caption$/i);
  return {
    headings,
    title: findStyle(styles, ["title"], /^title$/i),
    body: findStyle(styles, ["normal"], /^normal$/i) ?? findStyle(styles, ["body text"], /^bodytext$/i),
    listItem: findStyle(styles, ["list paragraph"], /^listparagraph$/i),
    figureCaption: findStyle(styles, ["figure caption"], /^figurecaption$/i) ?? caption,
    tableCaption: findStyle(styles, ["table caption"], /^tablecaption$/i) ?? caption,
    quote: findStyle(styles, ["quote"], /^quote$/i),
    code: findStyle(styles, ["source code", "code", "html preformatted"], /^(sourcecode|code|htmlpreformatted)$/i),
  };
}

/** Manual tag wins. A manual "heading" without a level uses the detected level, else 1. */
export function effectiveRole(detected: DetectedParagraph): EffectiveRole {
  const role = detected.manualRole ?? detected.role;
  if (role !== "heading") return { role, level: null };
  const raw = detected.level ?? 1;
  const level = Math.min(MAX_HEADING_LEVEL, Math.max(1, Math.round(raw)));
  return { role, level };
}

/** Deepest heading style at or above the wanted level; none if the template has no heading styles. */
export function headingStyleFor(map: StyleMap, level: number): string | null {
  const start = Math.min(MAX_HEADING_LEVEL, Math.max(1, Math.round(level)));
  for (let l = start; l >= 1; l -= 1) {
    const id = map.headings[l - 1];
    if (id !== undefined && id !== null) return id;
  }
  return null;
}

/** Style id for a role, or null when the paragraph must be left as it is. */
export function styleIdForRole(map: StyleMap, role: StructureRole, level: number | null): string | null {
  switch (role) {
    case "heading":
      return headingStyleFor(map, level ?? 1);
    case "title":
      return map.title;
    case "body":
      return map.body;
    case "list-item":
      return map.listItem;
    case "figure-caption":
      return map.figureCaption;
    case "table-caption":
      return map.tableCaption;
    case "quote":
      return map.quote;
    case "code":
      return map.code;
    case "table":
    case "unknown":
      return null;
  }
}

/** Style id for one detected paragraph (manual tag respected). */
export function styleIdForParagraph(map: StyleMap, detected: DetectedParagraph): string | null {
  const eff = effectiveRole(detected);
  return styleIdForRole(map, eff.role, eff.level);
}

/**
 * Names of roles that appear in the document but have no style in the
 * template, for the "template is missing styles" warning. Each name once,
 * for example "heading 3", "caption", "title".
 */
export function missingStyleNames(map: StyleMap, paragraphs: readonly DetectedParagraph[]): string[] {
  const missing = new Set<string>();
  for (const p of paragraphs) {
    const eff = effectiveRole(p);
    if (eff.role === "table" || eff.role === "unknown") continue;
    if (styleIdForRole(map, eff.role, eff.level) !== null) continue;
    if (eff.role === "heading") missing.add(`heading ${eff.level ?? 1}`);
    else if (eff.role === "figure-caption" || eff.role === "table-caption") missing.add("caption");
    else if (eff.role === "list-item") missing.add("list paragraph");
    else if (eff.role === "code") missing.add("source code");
    else missing.add(eff.role === "body" ? "normal" : eff.role);
  }
  return [...missing].sort();
}
