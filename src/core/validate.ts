/**
 * formatIQ output validation (file 3b-4).
 *
 * Checks a finished package for the mistakes that make Word say "the file is
 * corrupt" or silently lose things. It RETURNS a list of problems and never
 * throws for a bad package (it throws only for damaged XML, like the other
 * XML code).
 *
 * Two steps so big parts can be streamed:
 * 1. createPartScanner(name) is fed the XML of one part (write, end) and
 *    returns PartFacts (ids used, ids defined, section data).
 * 2. validatePackage(input) combines the facts with the part list, the .rels
 *    files and [Content_Types].xml.
 * scanXmlString is a shortcut for small parts and tests.
 *
 * Checks: required parts, duplicate part names (ignoring case), content type
 * for every part, Override for a missing part (warning), .rels well formed
 * (duplicate ids, internal targets that do not exist or climb out of the zip),
 * every r:id style attribute resolves in the rels of its own part, pStyle ids
 * exist in styles.xml (warning), numId exists in numbering.xml, bookmark ids
 * unique (names unique = warning), sectPr has a usable pgSz (warning) and no
 * repeated header/footer reference of one type.
 */
import { RawOpen, RawToken, createXmlRewriter } from "./xml_rewrite";
import { contentTypeOf, parseContentTypes, parseRelationships, partDir, relsPathOf, resolveTarget } from "./package_merge";

/** Namespace of r:id style attributes. */
export const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export type ProblemSeverity = "error" | "warning";

export interface ValidationProblem {
  severity: ProblemSeverity;
  /** Short stable code, e.g. "UNRESOLVED_RID". */
  code: string;
  /** Part the problem was found in. */
  part: string;
  message: string;
}

/** What one XML part uses and defines. */
export interface PartFacts {
  /** Values of attributes in the relationships namespace (r:id, r:embed...). */
  rIds: Set<string>;
  /** w:pStyle values. */
  pStyleIds: Set<string>;
  /** Non-zero w:numId values (as numbers). */
  numIds: Set<number>;
  /** w:styleId values defined by w:style elements. */
  styleDefIds: Set<string>;
  /** w:id of every bookmarkStart. */
  bookmarkIds: string[];
  /** w:name of every bookmarkStart. */
  bookmarkNames: string[];
  sectPrCount: number;
  /** Human readable section problems. */
  sectPrIssues: string[];
}

function localOf(name: string): string {
  const i = name.indexOf(":");
  return i < 0 ? name : name.slice(i + 1);
}

function prefixOf(name: string): string {
  const i = name.indexOf(":");
  return i < 0 ? "" : name.slice(0, i);
}

function attrLocal(tag: RawOpen, local: string): string | undefined {
  for (const [key, value] of tag.attrs) {
    if (!key.startsWith("xmlns") && localOf(key) === local) return value;
  }
  return undefined;
}

interface SectState { hasPgSz: boolean; badPgSz: boolean; refs: Map<string, number>; }

/** Feed one XML part in pieces; end() returns the facts. */
export function createPartScanner(partName: string): { write(text: string): void; end(): PartFacts } {
  const facts: PartFacts = {
    rIds: new Set(), pStyleIds: new Set(), numIds: new Set(), styleDefIds: new Set(),
    bookmarkIds: [], bookmarkNames: [], sectPrCount: 0, sectPrIssues: [],
  };
  const prefixes = new Map<string, string>();
  const sectStack: SectState[] = [];
  const finishSect = (state: SectState): void => {
    if (!state.hasPgSz) facts.sectPrIssues.push("a section has no page size (w:pgSz)");
    else if (state.badPgSz) facts.sectPrIssues.push("a section has a page size that is not a positive number");
    for (const [key, count] of state.refs) {
      if (count > 1) facts.sectPrIssues.push(`a section has ${count} header/footer references of type ${key}`);
    }
  };
  const handle = (token: RawToken): void => {
    if (token.kind === "open") {
      for (const [key, value] of token.attrs) {
        if (key.startsWith("xmlns:")) prefixes.set(key.slice(6), value);
      }
      for (const [key, value] of token.attrs) {
        if (key.startsWith("xmlns")) continue;
        const prefix = prefixOf(key);
        if (prefix !== "" && prefixes.get(prefix) === OFFICE_REL_NS) facts.rIds.add(value);
      }
      const local = localOf(token.name);
      if (local === "pStyle") {
        const v = attrLocal(token, "val");
        if (v !== undefined) facts.pStyleIds.add(v);
      } else if (local === "numId") {
        const v = attrLocal(token, "val");
        const n = v === undefined || !/^\d+$/.test(v.trim()) ? undefined : Number.parseInt(v.trim(), 10);
        if (n !== undefined && n !== 0) facts.numIds.add(n);
      } else if (local === "style") {
        const v = attrLocal(token, "styleId");
        if (v !== undefined) facts.styleDefIds.add(v);
      } else if (local === "bookmarkStart") {
        const id = attrLocal(token, "id");
        const name = attrLocal(token, "name");
        if (id !== undefined) facts.bookmarkIds.push(id);
        if (name !== undefined) facts.bookmarkNames.push(name);
      } else if (local === "sectPr") {
        facts.sectPrCount += 1;
        sectStack.push({ hasPgSz: false, badPgSz: false, refs: new Map() });
      } else if (local === "pgSz") {
        const top = sectStack[sectStack.length - 1];
        if (top !== undefined) {
          top.hasPgSz = true;
          const w = Number(attrLocal(token, "w"));
          const h = Number(attrLocal(token, "h"));
          if (!(w > 0) || !(h > 0)) top.badPgSz = true;
        }
      } else if (local === "headerReference" || local === "footerReference") {
        const top = sectStack[sectStack.length - 1];
        if (top !== undefined) {
          const key = local + ":" + (attrLocal(token, "type") ?? "default");
          top.refs.set(key, (top.refs.get(key) ?? 0) + 1);
        }
      }
    } else if (token.kind === "close" && localOf(token.name) === "sectPr") {
      const state = sectStack.pop();
      if (state !== undefined) finishSect(state);
    }
  };
  const rewriter = createXmlRewriter({ partName, onOutput: () => undefined, transform: (token) => handle(token) });
  return { write: (t) => rewriter.write(t), end: () => { rewriter.end(); return facts; } };
}

/** Scan a whole XML string. */
export function scanXmlString(partName: string, xml: string): PartFacts {
  const scanner = createPartScanner(partName);
  scanner.write(xml);
  return scanner.end();
}

export interface ValidationInput {
  /** Every part name in the output zip. */
  partNames: readonly string[];
  /** Text of every .rels file, keyed by its full name. */
  relsXml: ReadonlyMap<string, string>;
  /** Text of [Content_Types].xml (undefined if missing). */
  contentTypesXml: string | undefined;
  /** Facts of every XML part that was scanned, keyed by part name. */
  facts: ReadonlyMap<string, PartFacts>;
  /** Style ids defined in word/styles.xml (undefined when there is no styles part). */
  styleIds: ReadonlySet<string> | undefined;
  /** numIds defined in word/numbering.xml (undefined when there is no numbering part). */
  numberingIds: ReadonlySet<number> | undefined;
}

const MAIN = "word/document.xml";

/** Combine everything into a list of problems (empty list = clean). */
export function validatePackage(input: ValidationInput): ValidationProblem[] {
  const problems: ValidationProblem[] = [];
  const add = (severity: ProblemSeverity, code: string, part: string, message: string): void => {
    problems.push({ severity, code, part, message });
  };
  const names = new Set(input.partNames);

  const lower = new Map<string, string>();
  for (const name of input.partNames) {
    const key = name.toLowerCase();
    const other = lower.get(key);
    if (other !== undefined && other !== name) add("error", "DUPLICATE_PART", name, `Part names "${other}" and "${name}" differ only by case.`);
    lower.set(key, name);
  }
  if (!names.has(MAIN)) add("error", "MISSING_MAIN", MAIN, "The package has no word/document.xml.");
  if (!names.has("_rels/.rels")) add("error", "MISSING_ROOT_RELS", "_rels/.rels", "The package has no _rels/.rels.");

  // Content types.
  if (input.contentTypesXml === undefined) {
    add("error", "MISSING_CONTENT_TYPES", "[Content_Types].xml", "The package has no [Content_Types].xml.");
  } else {
    const ct = parseContentTypes(input.contentTypesXml);
    for (const name of input.partNames) {
      if (name === "[Content_Types].xml" || name.endsWith("/")) continue;
      if (contentTypeOf(ct, name) === undefined) add("error", "NO_CONTENT_TYPE", name, "This part has no content type.");
    }
    for (const key of ct.overrides.keys()) {
      if (!names.has(key.startsWith("/") ? key.slice(1) : key)) add("warning", "OVERRIDE_WITHOUT_PART", key, "[Content_Types].xml names a part that is not in the file.");
    }
  }

  // Relationship files.
  const idsByOwner = new Map<string, Set<string>>();
  for (const [relsName, xml] of input.relsXml) {
    const rels = parseRelationships(xml, relsName);
    const baseDir = partDir(partDir(relsName));
    const marker = "_rels/";
    const slash = relsName.lastIndexOf(marker);
    const ownerBase = slash < 0 ? "" : relsName.slice(slash + marker.length).replace(/\.rels$/, "");
    const owner = ownerBase === "" ? "" : (baseDir === "" ? "" : baseDir + "/") + ownerBase;
    const seen = new Set<string>();
    for (const rel of rels) {
      if (seen.has(rel.id)) add("error", "DUPLICATE_REL_ID", relsName, `Relationship id ${rel.id} is used twice.`);
      seen.add(rel.id);
      if (rel.targetMode === "External") continue;
      const target = resolveTarget(baseDir, rel.target);
      if (target === undefined) add("error", "REL_TARGET_OUTSIDE", relsName, `Relationship ${rel.id} points outside the package (${rel.target}).`);
      else if (!names.has(target)) add("error", "MISSING_TARGET", relsName, `Relationship ${rel.id} points to ${target}, which is not in the file.`);
    }
    idsByOwner.set(owner, seen);
  }

  // Facts per part.
  for (const [part, f] of input.facts) {
    const have = idsByOwner.get(part);
    for (const id of f.rIds) {
      if (have === undefined) {
        add("error", "UNRESOLVED_RID", part, `Uses relationship id ${id} but the part has no relationships file (${relsPathOf(part)}).`);
      } else if (!have.has(id)) {
        add("error", "UNRESOLVED_RID", part, `Uses relationship id ${id}, which is not in ${relsPathOf(part)}.`);
      }
    }
    if (input.styleIds !== undefined) {
      for (const id of f.pStyleIds) {
        if (!input.styleIds.has(id)) add("warning", "UNKNOWN_STYLE", part, `Paragraph style ${id} is not in styles.xml.`);
      }
    }
    for (const n of f.numIds) {
      if (input.numberingIds === undefined || !input.numberingIds.has(n)) add("error", "UNKNOWN_NUMID", part, `List id ${n} is not in numbering.xml.`);
    }
    const seenIds = new Set<string>();
    for (const id of f.bookmarkIds) {
      if (seenIds.has(id)) add("error", "DUPLICATE_BOOKMARK_ID", part, `Bookmark id ${id} is used twice.`);
      seenIds.add(id);
    }
    const seenNames = new Set<string>();
    for (const n of f.bookmarkNames) {
      if (seenNames.has(n)) add("warning", "DUPLICATE_BOOKMARK_NAME", part, `Bookmark name ${n} is used twice.`);
      seenNames.add(n);
    }
    for (const issue of f.sectPrIssues) add("warning", "BAD_SECTPR", part, issue);
  }
  return problems;
}

/** True when the list has no errors (warnings are allowed). */
export function hasErrors(problems: readonly ValidationProblem[]): boolean {
  return problems.some((p) => p.severity === "error");
}
