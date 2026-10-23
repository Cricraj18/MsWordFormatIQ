/**
 * formatIQ package merge (file 3b-2b).
 *
 * Pure functions on strings and small objects, no zip access. They decide
 * how the TEMPLATE package and the SOURCE document are combined:
 * - relationships (word/_rels/document.xml.rels),
 * - which source parts must be copied (with their own relationships),
 * - new names for source parts that clash with template parts,
 * - [Content_Types].xml.
 * apply.ts (next file) does the reading and writing.
 *
 * Rules (decided in chat 10, tested):
 * - Template-owned relationship types come from the TEMPLATE only: styles,
 *   theme, fontTable, settings, webSettings, numbering, header, footer.
 *   Source relationships of these types are dropped (and listed).
 * - Everything else stays from the SOURCE (images, hyperlinks, footnotes,
 *   endnotes, comments, charts, diagrams, embeddings, custom XML).
 * - Template relationships get NEW ids ("tpl" + old id) so they never
 *   collide with the ids that the source document.xml uses.
 * - A source part whose name is already used by the template gets a new name
 *   ("src_" + file name, then "src2_", "src3_"...). Its relationship target
 *   is fixed. Parts reached only through other source parts (for example an
 *   image used by footnotes.xml) are followed too.
 */
import { createXmlRewriter, escapeXmlAttr } from "./xml_rewrite";

/** Namespace of .rels files. */
export const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** Namespace of [Content_Types].xml. */
export const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/** One relationship. targetMode is "External" for web links and similar. */
export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/** Last words of the relationship types that the template owns. */
export const TEMPLATE_OWNED_SUFFIXES: readonly string[] = [
  "/styles",
  "/theme",
  "/fontTable",
  "/settings",
  "/webSettings",
  "/numbering",
  "/header",
  "/footer",
];

/** True when relationships of this type come from the template. */
export function isTemplateOwned(type: string): boolean {
  return TEMPLATE_OWNED_SUFFIXES.some((suffix) => type.endsWith(suffix));
}

function localOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function attrOf(attrs: readonly (readonly [string, string])[], name: string): string | undefined {
  for (const [key, value] of attrs) {
    if (localOf(key) === name) return value;
  }
  return undefined;
}

/** Read the elements called `element` of an XML string, giving their attributes. */
function readElements(xml: string, partName: string, element: string): (readonly [string, string])[][] {
  const found: (readonly [string, string])[][] = [];
  const rewriter = createXmlRewriter({
    partName,
    onOutput: () => undefined,
    transform: (token) => {
      if (token.kind === "open" && localOf(token.name) === element) found.push(token.attrs);
    },
  });
  rewriter.write(xml);
  rewriter.end();
  return found;
}

/** Read a .rels file. Elements without Id, Type or Target are skipped. */
export function parseRelationships(xml: string, partName = "rels"): Relationship[] {
  const list: Relationship[] = [];
  for (const attrs of readElements(xml, partName, "Relationship")) {
    const id = attrOf(attrs, "Id");
    const type = attrOf(attrs, "Type");
    const target = attrOf(attrs, "Target");
    if (id === undefined || type === undefined || target === undefined) continue;
    const mode = attrOf(attrs, "TargetMode");
    list.push(mode === undefined ? { id, type, target } : { id, type, target, targetMode: mode });
  }
  return list;
}

/** Write a .rels file. */
export function writeRelationships(rels: readonly Relationship[]): string {
  let out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELATIONSHIPS_NS}">`;
  for (const r of rels) {
    out += `<Relationship Id="${escapeXmlAttr(r.id)}" Type="${escapeXmlAttr(r.type)}" Target="${escapeXmlAttr(r.target)}"`;
    if (r.targetMode !== undefined) out += ` TargetMode="${escapeXmlAttr(r.targetMode)}"`;
    out += "/>";
  }
  return out + "</Relationships>";
}

/** Result of mergeRelationships. */
export interface MergedRelationships {
  /** Relationships for the output word/_rels/document.xml.rels. */
  rels: Relationship[];
  /** Template relationship id -> new id used in the output. */
  templateIdMap: Map<string, string>;
  /** Source relationships that were left out because the template owns that type. */
  droppedSourceRels: Relationship[];
}

/** Combine the relationships of the source and the template document parts. */
export function mergeRelationships(source: readonly Relationship[], template: readonly Relationship[]): MergedRelationships {
  const kept = source.filter((r) => !isTemplateOwned(r.type));
  const droppedSourceRels = source.filter((r) => isTemplateOwned(r.type));
  const used = new Set<string>(kept.map((r) => r.id));
  const templateIdMap = new Map<string, string>();
  const fromTemplate: Relationship[] = [];
  for (const r of template) {
    if (!isTemplateOwned(r.type)) continue;
    let id = "tpl" + r.id;
    let n = 2;
    while (used.has(id)) {
      id = "tpl" + n + "_" + r.id;
      n += 1;
    }
    used.add(id);
    templateIdMap.set(r.id, id);
    fromTemplate.push({ ...r, id });
  }
  return { rels: [...kept, ...fromTemplate], templateIdMap, droppedSourceRels };
}

// ---------------------------------------------------------------------------
// Part names and targets
// ---------------------------------------------------------------------------

/** Folder of a part name: "word/document.xml" -> "word", "[Content_Types].xml" -> "". */
export function partDir(partName: string): string {
  const slash = partName.lastIndexOf("/");
  return slash === -1 ? "" : partName.slice(0, slash);
}

function baseName(partName: string): string {
  const slash = partName.lastIndexOf("/");
  return slash === -1 ? partName : partName.slice(slash + 1);
}

/** Name of the .rels file of a part: "word/footnotes.xml" -> "word/_rels/footnotes.xml.rels". */
export function relsPathOf(partName: string): string {
  const dir = partDir(partName);
  return (dir === "" ? "" : dir + "/") + "_rels/" + baseName(partName) + ".rels";
}

/**
 * Turn a relationship target into a full part name inside the zip.
 * baseDir is the folder of the part that owns the relationship ("word").
 * A target that starts with "/" is already a full name. Segments "." and ".."
 * are resolved. Returns undefined if ".." climbs above the zip root.
 */
export function resolveTarget(baseDir: string, target: string): string | undefined {
  const parts: string[] = target.startsWith("/") ? [] : baseDir === "" ? [] : baseDir.split("/");
  const wanted = target.startsWith("/") ? target.slice(1) : target;
  for (const segment of wanted.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      if (parts.length === 0) return undefined;
      parts.pop();
    } else {
      parts.push(segment);
    }
  }
  return parts.join("/");
}

/** The shortest relative target from a folder to a full part name. */
export function relativeTarget(fromDir: string, partName: string): string {
  const from = fromDir === "" ? [] : fromDir.split("/");
  const to = partName.split("/");
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common += 1;
  const ups = from.length - common;
  return [...Array<string>(ups).fill(".."), ...to.slice(common)].join("/");
}

/**
 * All parts reached from the given relationships, following the relationships
 * of each reached part too (for example footnotes.xml -> an image).
 * relsByPart is keyed by full part name. External targets and targets that
 * leave the zip are ignored. Each part appears once, in discovery order.
 */
export function collectPartClosure(
  rootRels: readonly Relationship[],
  baseDir: string,
  relsByPart: ReadonlyMap<string, readonly Relationship[]>,
): string[] {
  const order: string[] = [];
  const seen = new Set<string>();
  const queue: { dir: string; rel: Relationship }[] = rootRels.map((rel) => ({ dir: baseDir, rel }));
  for (let i = 0; i < queue.length; i += 1) {
    const item = queue[i];
    if (item === undefined || item.rel.targetMode === "External") continue;
    const name = resolveTarget(item.dir, item.rel.target);
    if (name === undefined || name === "" || seen.has(name)) continue;
    seen.add(name);
    order.push(name);
    const children = relsByPart.get(name);
    if (children !== undefined) {
      for (const rel of children) queue.push({ dir: partDir(name), rel });
    }
  }
  return order;
}

/**
 * New names for source parts that clash with names used by the template.
 * Returns original name -> output name for EVERY part in sourceParts
 * (unchanged when there is no clash). Names are compared ignoring case,
 * because Windows and Word treat part names that way.
 */
export function planPartNames(sourceParts: readonly string[], templateParts: Iterable<string>): Map<string, string> {
  const taken = new Set<string>();
  for (const name of templateParts) taken.add(name.toLowerCase());
  const plan = new Map<string, string>();
  for (const original of sourceParts) {
    let output = original;
    let n = 1;
    while (taken.has(output.toLowerCase())) {
      const dir = partDir(original);
      const prefix = n === 1 ? "src_" : "src" + n + "_";
      output = (dir === "" ? "" : dir + "/") + prefix + baseName(original);
      n += 1;
    }
    taken.add(output.toLowerCase());
    plan.set(original, output);
  }
  return plan;
}

/**
 * Fix the targets of the relationships of one part after parts were renamed.
 * originalPart and newPart are the part's own full names (before and after).
 * renames is the map from planPartNames. Targets of renamed parts change,
 * every target is written relative to the part's NEW folder.
 */
export function retargetRels(
  rels: readonly Relationship[],
  originalPart: string,
  newPart: string,
  renames: ReadonlyMap<string, string>,
): Relationship[] {
  return rels.map((rel) => {
    if (rel.targetMode === "External") return rel;
    const abs = resolveTarget(partDir(originalPart), rel.target);
    if (abs === undefined) return rel;
    const moved = renames.get(abs) ?? abs;
    return { ...rel, target: relativeTarget(partDir(newPart), moved) };
  });
}

// ---------------------------------------------------------------------------
// [Content_Types].xml
// ---------------------------------------------------------------------------

/** Content types. defaults: extension (lower case) -> type. overrides: "/word/x.xml" -> type. */
export interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

/** Read [Content_Types].xml. */
export function parseContentTypes(xml: string): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const attrs of readElements(xml, "[Content_Types].xml", "Default")) {
    const ext = attrOf(attrs, "Extension");
    const type = attrOf(attrs, "ContentType");
    if (ext !== undefined && type !== undefined) defaults.set(ext.toLowerCase(), type);
  }
  for (const attrs of readElements(xml, "[Content_Types].xml", "Override")) {
    const name = attrOf(attrs, "PartName");
    const type = attrOf(attrs, "ContentType");
    if (name !== undefined && type !== undefined) overrides.set(name, type);
  }
  return { defaults, overrides };
}

/** Write [Content_Types].xml. */
export function writeContentTypes(ct: ContentTypes): string {
  let out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CONTENT_TYPES_NS}">`;
  for (const [ext, type] of ct.defaults) {
    out += `<Default Extension="${escapeXmlAttr(ext)}" ContentType="${escapeXmlAttr(type)}"/>`;
  }
  for (const [name, type] of ct.overrides) {
    out += `<Override PartName="${escapeXmlAttr(name)}" ContentType="${escapeXmlAttr(type)}"/>`;
  }
  return out + "</Types>";
}

function extensionOf(partName: string): string {
  const base = baseName(partName);
  const dot = base.lastIndexOf(".");
  return dot === -1 ? "" : base.slice(dot + 1).toLowerCase();
}

/** Content type of a part: its Override, else the Default of its extension, else undefined. */
export function contentTypeOf(ct: ContentTypes, partName: string): string | undefined {
  return ct.overrides.get("/" + partName) ?? ct.defaults.get(extensionOf(partName));
}

/**
 * Content types for the output. Starts from the template's list and adds what
 * the copied source parts need. copiedSourceParts maps the part's ORIGINAL
 * name to its OUTPUT name. A source part keeps its own content type: if the
 * template already gives that type to the output name nothing is added; if
 * the extension has no Default a Default is added; otherwise an Override.
 */
export function mergeContentTypes(
  template: ContentTypes,
  source: ContentTypes,
  copiedSourceParts: ReadonlyMap<string, string>,
): ContentTypes {
  const merged: ContentTypes = { defaults: new Map(template.defaults), overrides: new Map(template.overrides) };
  for (const [original, output] of copiedSourceParts) {
    const wanted = contentTypeOf(source, original);
    if (wanted === undefined) continue;
    if (contentTypeOf(merged, output) === wanted) continue;
    const ext = extensionOf(output);
    if (ext !== "" && !merged.defaults.has(ext) && !merged.overrides.has("/" + output)) {
      merged.defaults.set(ext, wanted);
    } else {
      merged.overrides.set("/" + output, wanted);
    }
  }
  return merged;
}

/** Remove Overrides for parts that are not in the output (the zip must not name missing parts). */
export function pruneOverrides(ct: ContentTypes, existingParts: Iterable<string>): ContentTypes {
  const exists = new Set<string>();
  for (const name of existingParts) exists.add("/" + name);
  const overrides = new Map<string, string>();
  for (const [name, type] of ct.overrides) {
    if (exists.has(name)) overrides.set(name, type);
  }
  return { defaults: new Map(ct.defaults), overrides };
}
