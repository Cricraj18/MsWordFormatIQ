/**
 * formatIQ numbering merge (file 3b-3).
 *
 * The output package uses the TEMPLATE's numbering.xml, but the source
 * document.xml points at list ids (w:numId) of the SOURCE numbering.xml.
 * This file builds one merged numbering.xml and a map from every source
 * numId to its new id, then rewrites w:numId in any XML part.
 *
 * Merge rules (all tested):
 * - Template entries are kept byte for byte and keep their ids.
 * - Source abstractNum ids are shifted by (largest template abstractNum id + 1),
 *   source num ids by the largest template num id, picture bullet ids by
 *   (largest template picture bullet id + 1). With no template entries the
 *   shift is 0 (source ids are kept).
 * - Source w:pStyle links inside list levels are removed (the source styles are
 *   not in the output package, the template styles are used instead).
 * - numId 0 means "no list" and is never changed.
 * - Source numIdMacAtCleanup is dropped. Template one is kept.
 * - Namespace declarations of the source root that the template root lacks are
 *   added to the output root (mc:Ignorable lists are united).
 *
 * Pure string functions. Throws FormatIQError for damaged XML or a DOCTYPE
 * (from xml_rewrite.ts).
 */
import {
  RawAttr,
  RawOpen,
  RawToken,
  RawTransform,
  XmlSerializer,
  createXmlRewriter,
  rewriteXmlString,
} from "./xml_rewrite";

/** Local part of a qualified name: "w:num" -> "num". */
function localOf(name: string): string {
  const i = name.indexOf(":");
  return i < 0 ? name : name.slice(i + 1);
}

/** Attribute value by LOCAL name (ignores xmlns attributes). */
function attrByLocal(tag: RawOpen, local: string): string | undefined {
  for (const [key, value] of tag.attrs) {
    if (!key.startsWith("xmlns") && localOf(key) === local) return value;
  }
  return undefined;
}

/** Copy of the tag with the attribute (found by local name) changed. */
function setAttrByLocal(tag: RawOpen, local: string, value: string): RawOpen {
  const attrs: RawAttr[] = tag.attrs.map((pair): RawAttr =>
    !pair[0].startsWith("xmlns") && localOf(pair[0]) === local ? [pair[0], value] : pair,
  );
  return { kind: "open", name: tag.name, attrs };
}

function toInt(value: string | undefined): number | undefined {
  if (value === undefined || !/^-?\d+$/.test(value.trim())) return undefined;
  return Number.parseInt(value.trim(), 10);
}

export type NumberingChildKind = "numPicBullet" | "abstractNum" | "num" | "other";

/** One top-level child of w:numbering, as raw XML text. */
export interface NumberingChild {
  kind: NumberingChildKind;
  /** numPicBulletId, abstractNumId or numId. */
  id: number | undefined;
  /** For kind "num": the abstractNumId it points to. */
  abstractNumId: number | undefined;
  xml: string;
}

export interface ParsedNumbering {
  rootOpen: RawOpen | undefined;
  children: NumberingChild[];
}

/** Split a numbering.xml into its root tag and top-level children. */
export function parseNumbering(xml: string, partName = "word/numbering.xml"): ParsedNumbering {
  const children: NumberingChild[] = [];
  let rootOpen: RawOpen | undefined;
  let depth = 0;
  let serializer: XmlSerializer | undefined;
  let pieces: string[] = [];
  let current: NumberingChild | undefined;
  const transform: RawTransform = (token: RawToken): void => {
    if (token.kind === "open") {
      depth += 1;
      if (depth === 1) rootOpen = token;
      if (depth === 2) {
        const local = localOf(token.name);
        pieces = [];
        serializer = new XmlSerializer((chunk) => pieces.push(chunk));
        const kind: NumberingChildKind =
          local === "numPicBullet" || local === "abstractNum" || local === "num" ? local : "other";
        const idName = kind === "numPicBullet" ? "numPicBulletId" : kind === "abstractNum" ? "abstractNumId" : "numId";
        current = { kind, id: kind === "other" ? undefined : toInt(attrByLocal(token, idName)), abstractNumId: undefined, xml: "" };
      }
      if (depth === 3 && current?.kind === "num" && localOf(token.name) === "abstractNumId") {
        current.abstractNumId = toInt(attrByLocal(token, "val"));
      }
      if (depth >= 2) serializer?.push(token);
      return;
    }
    if (token.kind === "close") {
      if (depth >= 2) serializer?.push(token);
      if (depth === 2 && serializer !== undefined && current !== undefined) {
        serializer.flush();
        current.xml = pieces.join("");
        children.push(current);
        serializer = undefined;
        current = undefined;
      }
      depth -= 1;
      return;
    }
    if (depth >= 2) serializer?.push(token);
  };
  const rewriter = createXmlRewriter({ partName, onOutput: () => undefined, transform });
  rewriter.write(xml);
  rewriter.end();
  return { rootOpen, children };
}

/** All numIds defined in a numbering.xml (for validate.ts). */
export function listNumIds(numberingXml: string): Set<number> {
  const ids = new Set<number>();
  for (const child of parseNumbering(numberingXml).children) {
    if (child.kind === "num" && child.id !== undefined) ids.add(child.id);
  }
  return ids;
}

export interface MergedNumbering {
  /** The merged numbering.xml, or undefined when neither side has one. */
  xml: string | undefined;
  /** Source numId -> output numId. */
  numIdMap: Map<number, number>;
  abstractIdMap: Map<number, number>;
  picBulletIdMap: Map<number, number>;
  warnings: string[];
}

function maxId(children: NumberingChild[], kind: NumberingChildKind): number | undefined {
  let max: number | undefined;
  for (const c of children) {
    if (c.kind === kind && c.id !== undefined && (max === undefined || c.id > max)) max = c.id;
  }
  return max;
}

interface ShiftContext {
  abstractShift: number;
  numShift: number;
  picShift: number;
  droppedPStyle: number;
  styleLinks: number;
}

/** Rewrite one source child with shifted ids. */
function shiftChild(child: NumberingChild, ctx: ShiftContext, partName: string): string {
  let skipDepth = 0;
  let depth = 0;
  const transform: RawTransform = (token, emit) => {
    if (token.kind === "open") {
      depth += 1;
      const local = localOf(token.name);
      if (skipDepth > 0) { skipDepth += 1; return; }
      if (depth === 1) {
        if (child.kind === "abstractNum") {
          const id = toInt(attrByLocal(token, "abstractNumId"));
          emit(id === undefined ? token : setAttrByLocal(token, "abstractNumId", String(id + ctx.abstractShift)));
          return;
        }
        if (child.kind === "num") {
          const id = toInt(attrByLocal(token, "numId"));
          emit(id === undefined ? token : setAttrByLocal(token, "numId", String(id + ctx.numShift)));
          return;
        }
        if (child.kind === "numPicBullet") {
          const id = toInt(attrByLocal(token, "numPicBulletId"));
          emit(id === undefined ? token : setAttrByLocal(token, "numPicBulletId", String(id + ctx.picShift)));
          return;
        }
      }
      if (child.kind === "abstractNum" && local === "pStyle") { skipDepth = 1; ctx.droppedPStyle += 1; return; }
      if (child.kind === "abstractNum" && (local === "styleLink" || local === "numStyleLink")) ctx.styleLinks += 1;
      if (child.kind === "abstractNum" && local === "lvlPicBulletId") {
        const id = toInt(attrByLocal(token, "val"));
        emit(id === undefined ? token : setAttrByLocal(token, "val", String(id + ctx.picShift)));
        return;
      }
      if (child.kind === "num" && local === "abstractNumId" && depth === 2) {
        const id = toInt(attrByLocal(token, "val"));
        emit(id === undefined ? token : setAttrByLocal(token, "val", String(id + ctx.abstractShift)));
        return;
      }
      emit(token);
      return;
    }
    if (token.kind === "close") {
      depth -= 1;
      if (skipDepth > 0) { skipDepth -= 1; return; }
      emit(token);
      return;
    }
    if (skipDepth > 0) return;
    emit(token);
  };
  return rewriteXmlString(child.xml, partName, transform);
}

/** Union of namespace declarations and mc:Ignorable of two roots. */
function mergeRoot(template: RawOpen | undefined, source: RawOpen | undefined, warnings: string[]): RawOpen {
  const base: RawOpen = template ?? source ?? { kind: "open", name: "w:numbering", attrs: [] };
  if (template === undefined || source === undefined) return base;
  const attrs: RawAttr[] = base.attrs.map((p): RawAttr => p);
  for (const [key, value] of source.attrs) {
    if (key.startsWith("xmlns")) {
      const existing = attrs.find((a) => a[0] === key);
      if (existing === undefined) attrs.push([key, value]);
      else if (existing[1] !== value) warnings.push(`Namespace ${key} differs between template and source numbering.`);
    } else if (localOf(key) === "Ignorable") {
      const index = attrs.findIndex((a) => a[0] === key);
      const have = index >= 0 ? (attrs[index]?.[1] ?? "").split(/\s+/).filter((s) => s !== "") : [];
      for (const word of value.split(/\s+/).filter((s) => s !== "")) if (!have.includes(word)) have.push(word);
      if (index >= 0) attrs[index] = [key, have.join(" ")];
      else attrs.push([key, have.join(" ")]);
    }
  }
  return { kind: "open", name: base.name, attrs };
}

/**
 * Merge template and source numbering.xml text (either can be undefined).
 * Never mutates its inputs. Same inputs give the same output.
 */
export function mergeNumbering(
  templateXml: string | undefined,
  sourceXml: string | undefined,
  sourcePartName = "word/numbering.xml",
): MergedNumbering {
  const numIdMap = new Map<number, number>();
  const abstractIdMap = new Map<number, number>();
  const picBulletIdMap = new Map<number, number>();
  const warnings: string[] = [];
  if (sourceXml === undefined) return { xml: templateXml, numIdMap, abstractIdMap, picBulletIdMap, warnings };

  const template = templateXml === undefined ? undefined : parseNumbering(templateXml, "template numbering.xml");
  const source = parseNumbering(sourceXml, sourcePartName);
  const tChildren = template?.children ?? [];
  const tAbs = maxId(tChildren, "abstractNum");
  const tNum = maxId(tChildren, "num");
  const tPic = maxId(tChildren, "numPicBullet");
  const ctx: ShiftContext = {
    abstractShift: tAbs === undefined ? 0 : tAbs + 1,
    numShift: tNum === undefined ? 0 : tNum,
    picShift: tPic === undefined ? 0 : tPic + 1,
    droppedPStyle: 0,
    styleLinks: 0,
  };

  const bullets: string[] = [];
  const abstracts: string[] = [];
  const nums: string[] = [];
  const others: string[] = [];
  for (const c of tChildren) {
    if (c.kind === "numPicBullet") bullets.push(c.xml);
    else if (c.kind === "abstractNum") abstracts.push(c.xml);
    else if (c.kind === "num") nums.push(c.xml);
    else others.push(c.xml);
  }
  for (const c of source.children) {
    if (c.kind === "other") continue;
    const xml = shiftChild(c, ctx, sourcePartName);
    if (c.kind === "numPicBullet") {
      bullets.push(xml);
      if (c.id !== undefined) picBulletIdMap.set(c.id, c.id + ctx.picShift);
    } else if (c.kind === "abstractNum") {
      abstracts.push(xml);
      if (c.id !== undefined) abstractIdMap.set(c.id, c.id + ctx.abstractShift);
    } else {
      nums.push(xml);
      if (c.id !== undefined) numIdMap.set(c.id, c.id + ctx.numShift);
    }
  }
  if (ctx.droppedPStyle > 0) warnings.push(`Removed ${ctx.droppedPStyle} style link(s) from source list levels (template styles are used).`);
  if (ctx.styleLinks > 0) warnings.push(`${ctx.styleLinks} source list definition(s) use a style link that may not resolve.`);

  const root = mergeRoot(template?.rootOpen, source.rootOpen, warnings);
  const out: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'];
  const serializer = new XmlSerializer((chunk) => out.push(chunk));
  serializer.push(root);
  for (const xml of [...bullets, ...abstracts, ...nums, ...others]) serializer.push({ kind: "raw", xml });
  serializer.push({ kind: "close", name: root.name });
  serializer.flush();
  return { xml: out.join(""), numIdMap, abstractIdMap, picBulletIdMap, warnings };
}

/** Run two transforms in a row: every token the first emits goes through the second. */
export function chainTransforms(first: RawTransform, second: RawTransform): RawTransform {
  return (token, emit) => first(token, (t) => second(t, emit));
}

export interface NumIdRemapStats {
  remapped: number;
  /** w:numId values (not 0) that are not in the map and were left alone. */
  unmapped: number;
}

/** Transform that rewrites the value of every w:numId element using the map. */
export function createNumIdRemapper(map: ReadonlyMap<number, number>): { transform: RawTransform; stats: NumIdRemapStats } {
  const stats: NumIdRemapStats = { remapped: 0, unmapped: 0 };
  const transform: RawTransform = (token, emit) => {
    if (token.kind === "open" && localOf(token.name) === "numId") {
      const value = toInt(attrByLocal(token, "val"));
      if (value !== undefined && value !== 0) {
        const next = map.get(value);
        if (next === undefined) {
          stats.unmapped += 1;
        } else {
          stats.remapped += 1;
          emit(setAttrByLocal(token, "val", String(next)));
          return;
        }
      }
    }
    emit(token);
  };
  return { transform, stats };
}

/** Rewrite w:numId in a whole XML text (for copied source parts such as footnotes). */
export function remapNumIds(xml: string, partName: string, map: ReadonlyMap<number, number>): { xml: string; stats: NumIdRemapStats } {
  const remapper = createNumIdRemapper(map);
  return { xml: rewriteXmlString(xml, partName, remapper.transform), stats: remapper.stats };
}
