/**
 * formatIQ watermark (file 5c, version 2: shared types come from ../types).
 *
 * A Word watermark is a VML shape inside a run of a HEADER part
 * (word/header1.xml ...): <w:r><w:pict><v:shape id="PowerPlusWaterMarkObject..."
 * ...> for text and id="WordPictureWatermark..." for a picture. Every header
 * type that is used (default, first page, even pages) needs its own copy.
 *
 * Pure functions on the XML text of ONE header part. No zip access: apply.ts
 * calls setWatermark for each header part it writes.
 *
 * What is done (tested):
 * - findWatermarks: lists watermark shapes of a header (kind, text, colour...).
 * - removeWatermarks: drops the whole run that holds a watermark shape.
 * - setWatermark(mode): "inherit" and "keepSource" leave the header as it is;
 *   "remove" removes; "text" and "image" remove an old one first and add the
 *   new one, so a header never ends up with two watermarks.
 * - The new run goes into the FIRST paragraph of the header (after its
 *   w:pPr). A header without a paragraph gets one. Missing namespace
 *   declarations (v, o, w10, r) are added to the root tag.
 *
 * UNSURE / NOT SEEN IN WORD: the VML (shape types, formulas, sizes, washout
 * values) is written from knowledge of what Word itself writes. It is
 * well-formed XML and re-reads correctly in tests, but how Word draws it has
 * not been checked. Sizes are a simple estimate from the text length.
 * Standard prefixes w:, v:, o:, w10:, r: are assumed.
 */
import { FormatIQError } from "../types";
import type { WatermarkMode, WatermarkOptions } from "../types";
import { closeTag, createXmlRewriter, openTag, rewriteXmlString } from "./xml_rewrite";
import type { RawAttr, RawOpen, RawToken, RawTransform } from "./xml_rewrite";
import { parseRelationships, writeRelationships } from "./package_merge";
import type { Relationship } from "./package_merge";

export type { WatermarkMode, WatermarkOptions };

/** Word names text watermark shapes with this word in their id. */
export const TEXT_MARKER = "PowerPlusWaterMarkObject";

/** Word names picture watermark shapes with this word in their id. */
export const IMAGE_MARKER = "WordPictureWatermark";

/** One watermark found in a header. */
export interface WatermarkInfo {
  kind: "text" | "image";
  id: string;
  /** Text of a text watermark. */
  text: string | undefined;
  /** Fill colour as written (for example "silver" or "#C0C0C0"). */
  color: string | undefined;
  /** Opacity as written (for example ".5"). */
  opacity: string | undefined;
  /** True when the style has a rotation. */
  rotated: boolean;
  /** Picture relationship id of an image watermark. */
  relId: string | undefined;
}

function localOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function attrByLocal(tag: RawOpen, local: string): string | undefined {
  for (const [key, value] of tag.attrs) {
    if (!key.startsWith("xmlns") && localOf(key) === local) return value;
  }
  return undefined;
}

function markerKind(id: string | undefined): "text" | "image" | undefined {
  if (id === undefined) return undefined;
  if (id.includes(TEXT_MARKER)) return "text";
  if (id.includes(IMAGE_MARKER)) return "image";
  return undefined;
}

function invalid(message: string, detail?: string): FormatIQError {
  return new FormatIQError("INVALID_OPTIONS", message, detail);
}

// ---------------------------------------------------------------------------
// Finding and removing
// ---------------------------------------------------------------------------

/** List the watermark shapes of a header part. */
export function findWatermarks(xml: string, partName = "word/header1.xml"): WatermarkInfo[] {
  const found: WatermarkInfo[] = [];
  let current: WatermarkInfo | undefined;
  let depth = 0;
  let shapeDepth = 0;
  const rewriter = createXmlRewriter({
    partName,
    onOutput: () => undefined,
    transform: (token) => {
      if (token.kind === "open") {
        depth += 1;
        const local = localOf(token.name);
        if (local === "shape" && current === undefined) {
          const id = attrByLocal(token, "id");
          const kind = markerKind(id);
          if (kind !== undefined && id !== undefined) {
            const style = attrByLocal(token, "style") ?? "";
            current = { kind, id, text: undefined, color: attrByLocal(token, "fillcolor"), opacity: undefined, rotated: /(^|;)\s*rotation\s*:/.test(style), relId: undefined };
            shapeDepth = depth;
          }
        } else if (current !== undefined) {
          if (local === "textpath") current.text = attrByLocal(token, "string");
          else if (local === "fill") current.opacity = attrByLocal(token, "opacity");
          else if (local === "imagedata") current.relId = attrByLocal(token, "id");
        }
      } else if (token.kind === "close") {
        if (current !== undefined && depth === shapeDepth) {
          found.push(current);
          current = undefined;
        }
        depth -= 1;
      }
    },
  });
  rewriter.write(xml);
  rewriter.end();
  return found;
}

/** Transform that drops every run (w:r) that holds a watermark shape. */
export function createWatermarkRemover(): { transform: RawTransform; stats: { removed: number } } {
  const stats = { removed: 0 };
  let buffer: RawToken[] | null = null;
  let depth = 0;
  let hasMark = false;
  const transform: RawTransform = (token, emit) => {
    if (buffer !== null) {
      buffer.push(token);
      if (token.kind === "open") {
        depth += 1;
        if (localOf(token.name) === "shape" && markerKind(attrByLocal(token, "id")) !== undefined) hasMark = true;
      } else if (token.kind === "close") {
        depth -= 1;
      }
      if (token.kind === "close" && depth === 0) {
        const done = buffer;
        buffer = null;
        if (hasMark) stats.removed += 1;
        else for (const t of done) emit(t);
      }
      return;
    }
    if (token.kind === "open" && localOf(token.name) === "r") {
      buffer = [token];
      depth = 1;
      hasMark = false;
      return;
    }
    emit(token);
  };
  return { transform, stats };
}

/** Remove all watermarks from a header part. */
export function removeWatermarks(xml: string, partName = "word/header1.xml"): { xml: string; removed: number } {
  const remover = createWatermarkRemover();
  return { xml: rewriteXmlString(xml, partName, remover.transform), removed: remover.stats.removed };
}

// ---------------------------------------------------------------------------
// Building the new shape
// ---------------------------------------------------------------------------

interface Node {
  name: string;
  attrs?: RawAttr[];
  children?: Node[];
}

function nodeTokens(node: Node, out: RawToken[]): void {
  out.push(openTag(node.name, node.attrs ?? []));
  for (const child of node.children ?? []) nodeTokens(child, out);
  out.push(closeTag(node.name));
}

const leaf = (name: string, attrs: RawAttr[] = []): Node => ({ name, attrs });
const eqn = (value: string): Node => leaf("v:f", [["eqn", value]]);

const TEXT_SHAPETYPE: Node = {
  name: "v:shapetype",
  attrs: [
    ["id", "_x0000_t136"],
    ["coordsize", "21600,21600"],
    ["o:spt", "136"],
    ["adj", "10800"],
    ["path", "m@7,l@8,m@5,21600l@6,21600e"],
  ],
  children: [
    {
      name: "v:formulas",
      children: [
        "sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3", "if @0 @3 0", "if @0 21600 @1",
        "if @0 0 @2", "if @0 @4 21600", "mid @5 @6", "mid @8 @5", "mid @7 @8", "mid @6 @7", "sum @6 0 @5",
      ].map(eqn),
    },
    leaf("v:path", [["textpathok", "t"], ["o:connecttype", "custom"], ["o:connectlocs", "@9,0;@10,10800;@11,21600;@12,10800"], ["o:connectangles", "270,180,90,0"]]),
    leaf("v:textpath", [["on", "t"], ["fitshape", "t"]]),
    { name: "v:handles", children: [leaf("v:h", [["position", "#0,bottomRight"], ["xrange", "6629,14971"]])] },
    leaf("o:lock", [["v:ext", "edit"], ["text", "t"], ["shapetype", "t"]]),
  ],
};

const PICTURE_SHAPETYPE: Node = {
  name: "v:shapetype",
  attrs: [
    ["id", "_x0000_t75"],
    ["coordsize", "21600,21600"],
    ["o:spt", "75"],
    ["o:preferrelative", "t"],
    ["path", "m@4@5l@4@11@9@11@9@5xe"],
    ["filled", "f"],
    ["stroked", "f"],
  ],
  children: [
    leaf("v:stroke", [["joinstyle", "miter"]]),
    {
      name: "v:formulas",
      children: [
        "if lineDrawn pixelLineWidth 0", "sum @0 1 0", "sum 0 0 @1", "prod @2 1 2", "prod @3 21600 pixelWidth", "prod @3 21600 pixelHeight",
        "sum @0 0 1", "prod @6 1 2", "prod @7 21600 pixelWidth", "sum @8 21600 0", "prod @7 21600 pixelHeight", "sum @10 21600 0",
      ].map(eqn),
    },
    leaf("v:path", [["o:extrusionok", "f"], ["gradientshapeok", "t"], ["o:connecttype", "rect"]]),
    leaf("o:lock", [["v:ext", "edit"], ["aspectratio", "t"]]),
  ],
};

function shapeStyle(widthPt: number, heightPt: number, rotate: boolean, seed: number): string {
  const z = -251658752 - (seed - 1) * 1024;
  return (
    `position:absolute;margin-left:0;margin-top:0;width:${widthPt}pt;height:${heightPt}pt;` +
    (rotate ? "rotation:315;" : "") +
    `z-index:${z};mso-position-horizontal:center;mso-position-horizontal-relative:margin;` +
    "mso-position-vertical:center;mso-position-vertical-relative:margin"
  );
}

function checkSeed(seed: number | undefined): number {
  const value = seed ?? 1;
  if (!Number.isInteger(value) || value < 1 || value > 1_000_000) throw invalid("The watermark number must be a whole number from 1 to 1000000.");
  return value;
}

/** Opacity as Word writes it: .5 for 50 percent, 1 for solid. */
export function formatOpacity(opacity: number): string {
  const rounded = Math.round(opacity * 100) / 100;
  const text = String(rounded);
  return text.startsWith("0.") ? text.slice(1) : text;
}

/** Tokens of one run holding a text watermark. Throws INVALID_OPTIONS for bad values. */
export function buildTextWatermarkRun(options: WatermarkOptions): RawToken[] {
  const text = (options.text ?? "").trim();
  if (text.length === 0) throw invalid("Enter the watermark text.");
  if (text.length > 100) throw invalid("The watermark text can have at most 100 characters.");
  const color = options.color ?? "C0C0C0";
  if (!/^[0-9A-Fa-f]{6}$/.test(color)) throw invalid("The watermark colour must be six hex digits, for example C0C0C0.", color);
  const opacity = options.opacity ?? 0.5;
  if (!(opacity > 0 && opacity <= 1)) throw invalid("The watermark opacity must be more than 0 and at most 1.");
  const fontFamily = (options.fontFamily ?? "Calibri").trim();
  if (fontFamily.length === 0 || fontFamily.length > 64) throw invalid("The watermark font name is not valid.");
  const seed = checkSeed(options.seed);
  const diagonal = options.diagonal ?? true;
  const width = Math.min(520, Math.max(150, Math.round(text.length * 70)));
  const height = Math.round(width * (diagonal ? 0.6 : 0.25));
  const shape: Node = {
    name: "v:shape",
    attrs: [
      ["id", `${TEXT_MARKER}${100000000 + seed}`],
      ["o:spid", `_x0000_s${2048 + seed}`],
      ["type", "#_x0000_t136"],
      ["style", shapeStyle(width, height, diagonal, seed)],
      ["o:allowincell", "f"],
      ["fillcolor", "#" + color.toUpperCase()],
      ["stroked", "f"],
    ],
    children: [
      leaf("v:fill", [["opacity", formatOpacity(opacity)]]),
      leaf("v:textpath", [["style", `font-family:"${fontFamily}";font-size:1pt`], ["string", text]]),
      leaf("w10:wrap", [["anchorx", "margin"], ["anchory", "margin"]]),
    ],
  };
  const run: Node = {
    name: "w:r",
    children: [{ name: "w:rPr", children: [leaf("w:noProof")] }, { name: "w:pict", children: [TEXT_SHAPETYPE, shape] }],
  };
  const out: RawToken[] = [];
  nodeTokens(run, out);
  return out;
}

/** Tokens of one run holding a picture watermark. Throws INVALID_OPTIONS for bad values. */
export function buildImageWatermarkRun(options: WatermarkOptions): RawToken[] {
  const relId = options.imageRelId;
  if (relId === undefined || relId.trim() === "") throw invalid("The watermark picture is missing.");
  const width = options.widthPt;
  const height = options.heightPt;
  if (width === undefined || height === undefined || !(width > 0 && width <= 1500) || !(height > 0 && height <= 1500)) {
    throw invalid("The watermark picture needs a width and a height in points (up to 1500).");
  }
  const seed = checkSeed(options.seed);
  const washout = options.washout ?? true;
  const imageAttrs: RawAttr[] = [["r:id", relId], ["o:title", ""]];
  if (washout) imageAttrs.push(["gain", "19661f"], ["blacklevel", "22938f"]);
  const shape: Node = {
    name: "v:shape",
    attrs: [
      ["id", `${IMAGE_MARKER}${100000000 + seed}`],
      ["o:spid", `_x0000_s${2048 + seed}`],
      ["type", "#_x0000_t75"],
      ["style", shapeStyle(Math.round(width), Math.round(height), false, seed)],
      ["o:allowincell", "f"],
    ],
    children: [leaf("v:imagedata", imageAttrs), leaf("w10:wrap", [["anchorx", "margin"], ["anchory", "margin"]])],
  };
  const run: Node = {
    name: "w:r",
    children: [{ name: "w:rPr", children: [leaf("w:noProof")] }, { name: "w:pict", children: [PICTURE_SHAPETYPE, shape] }],
  };
  const out: RawToken[] = [];
  nodeTokens(run, out);
  return out;
}

// ---------------------------------------------------------------------------
// Putting the run into the header
// ---------------------------------------------------------------------------

const NAMESPACES: readonly RawAttr[] = [
  ["xmlns:v", "urn:schemas-microsoft-com:vml"],
  ["xmlns:o", "urn:schemas-microsoft-com:office:office"],
  ["xmlns:w10", "urn:schemas-microsoft-com:office:word"],
  ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
];

/**
 * Transform that puts a run into the first paragraph that is a direct child
 * of the root (after its w:pPr), or adds a paragraph at the end of the root
 * when there is none. Adds missing namespace declarations to the root tag.
 */
export function createRunInserter(run: readonly RawToken[]): { transform: RawTransform; stats: { inserted: boolean } } {
  const stats = { inserted: false };
  let depth = 0;
  let pending = false;
  let inPPr = false;
  let pPrDepth = 0;
  const transform: RawTransform = (token, emit) => {
    if (token.kind === "open") {
      depth += 1;
      if (depth === 1) {
        const have = new Set(token.attrs.map((a) => a[0]));
        const extra = NAMESPACES.filter((n) => !have.has(n[0]));
        emit(extra.length === 0 ? token : openTag(token.name, [...token.attrs, ...extra]));
        return;
      }
      if (pending && !inPPr) {
        if (localOf(token.name) === "pPr" && depth === 3) {
          inPPr = true;
          pPrDepth = depth;
          emit(token);
          return;
        }
        for (const t of run) emit(t);
        stats.inserted = true;
        pending = false;
      }
      if (!stats.inserted && !pending && depth === 2 && localOf(token.name) === "p") pending = true;
      emit(token);
      return;
    }
    if (token.kind === "close") {
      if (inPPr && depth === pPrDepth) {
        emit(token);
        inPPr = false;
        for (const t of run) emit(t);
        stats.inserted = true;
        pending = false;
        depth -= 1;
        return;
      }
      if (pending && depth === 2) {
        for (const t of run) emit(t);
        stats.inserted = true;
        pending = false;
      }
      if (depth === 1 && !stats.inserted) {
        emit(openTag("w:p"));
        for (const t of run) emit(t);
        emit(closeTag("w:p"));
        stats.inserted = true;
      }
      depth -= 1;
      emit(token);
      return;
    }
    if (pending && !inPPr && token.kind === "text" && token.text.trim() !== "") {
      for (const t of run) emit(t);
      stats.inserted = true;
      pending = false;
    }
    emit(token);
  };
  return { transform, stats };
}

/** Result of setWatermark. */
export interface SetWatermarkResult {
  xml: string;
  /** Watermark runs that were removed. */
  removed: number;
  /** True when a new watermark was added. */
  added: boolean;
}

/**
 * Apply the chosen mode to the XML of one header part.
 * "inherit" and "keepSource" return the text unchanged.
 * Throws FormatIQError INVALID_OPTIONS for bad values, ZIP_READ_FAILED for damaged XML.
 */
export function setWatermark(xml: string, partName: string, options: WatermarkOptions): SetWatermarkResult {
  if (options.mode === "inherit" || options.mode === "keepSource") return { xml, removed: 0, added: false };
  if (options.mode === "remove") {
    const r = removeWatermarks(xml, partName);
    return { xml: r.xml, removed: r.removed, added: false };
  }
  const run = options.mode === "text" ? buildTextWatermarkRun(options) : buildImageWatermarkRun(options);
  const cleaned = removeWatermarks(xml, partName);
  const inserter = createRunInserter(run);
  const out = rewriteXmlString(cleaned.xml, partName, inserter.transform);
  return { xml: out, removed: cleaned.removed, added: inserter.stats.inserted };
}

/**
 * Add a picture relationship to the text of a header's .rels file
 * (undefined = the header has none yet). Returns the new file text and the
 * new relationship id ("rIdWm1", "rIdWm2"... first free one).
 * The picture file itself and its content type are added by apply.ts.
 */
export function addImageRelationship(relsXml: string | undefined, target: string): { xml: string; id: string } {
  const rels: Relationship[] = relsXml === undefined ? [] : parseRelationships(relsXml, "header rels");
  const used = new Set(rels.map((r) => r.id));
  let n = 1;
  while (used.has("rIdWm" + n)) n += 1;
  const id = "rIdWm" + n;
  rels.push({ id, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target });
  return { xml: writeRelationships(rels), id };
}
