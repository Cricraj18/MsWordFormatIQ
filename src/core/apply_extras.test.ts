/**
 * Tests for apply.ts "version 2": page setup and watermark (HANDOFF_v33 REMAINS item 2).
 * Separate file so the tested apply.test.ts stays as it is.
 * STATUS: WRITTEN, NOT YET RUN. Expected values for page setup come from the PageSetupOptions
 * comments in types.ts; if a test fails, check first whether the test or apply.ts is wrong.
 */
import { BlobWriter, TextReader, ZipWriter } from "@zip.js/zip.js";
import { beforeAll, describe, expect, it } from "vitest";
import { createDefaultFormatOptions, isFormatIQError } from "../types";
import type { FormatOptions, PageSetupOptions, WatermarkOptions } from "../types";
import { applyTemplate } from "./apply";
import { openDocx, readEntryText, setWorkersEnabled } from "./zipio";
import type { OpenedDocx } from "./zipio";
import { parseRelationships } from "./package_merge";
import { scanXmlString, validatePackage } from "./validate";
import type { PartFacts, ValidationProblem } from "./validate";
import { setWatermark } from "./watermark";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"';
const PK = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";

beforeAll(() => { setWorkersEnabled(false); });

async function makeZip(parts: Record<string, string>): Promise<Blob> {
  const writer = new BlobWriter("application/zip");
  const zip = new ZipWriter(writer);
  for (const [name, text] of Object.entries(parts)) await zip.add(name, new TextReader(text));
  await zip.close();
  return writer.getData();
}

function rels(items: [string, string, string][]): string {
  const body = items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships ${PK}>${body}</Relationships>`;
}

function style(id: string, name: string): string {
  return `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/></w:style>`;
}

const STYLES = `<w:styles ${W}>${style("Normal", "Normal")}${style("Heading1", "heading 1")}</w:styles>`;
const HEADER_PLAIN = `<w:hdr ${W} ${R}><w:p><w:r><w:t>TEMPLATE HEADER</w:t></w:r></w:p></w:hdr>`;

interface TemplateOptions {
  /** false = the template has no header part at all. Default true. */
  header?: boolean;
  /** Text of a watermark already inside the template header. */
  templateWatermark?: string;
}

function templateParts(o: TemplateOptions = {}): Record<string, string> {
  const withHeader = o.header !== false;
  const sect = `<w:sectPr>${withHeader ? '<w:headerReference w:type="default" r:id="rId10"/>' : ""}<w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
  const docRels: [string, string, string][] = [["rId1", "styles", "styles.xml"], ["rId3", "settings", "settings.xml"]];
  if (withHeader) docRels.push(["rId10", "header", "header1.xml"]);
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types ${CT}><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${WML}.settings+xml"/>${withHeader ? `<Override PartName="/word/header1.xml" ContentType="${WML}.header+xml"/>` : ""}</Types>`,
    "_rels/.rels": rels([["rId1", "officeDocument", "word/document.xml"]]),
    "word/document.xml": `<?xml version="1.0"?><w:document ${W} ${R}><w:body><w:p><w:r><w:t>Template body text</w:t></w:r></w:p>${sect}</w:body></w:document>`,
    "word/_rels/document.xml.rels": rels(docRels),
    "word/styles.xml": STYLES,
    "word/settings.xml": `<w:settings ${W}><w:zoom w:percent="100"/><w:compat/></w:settings>`,
  };
  if (withHeader) {
    const header = o.templateWatermark === undefined
      ? HEADER_PLAIN
      : setWatermark(HEADER_PLAIN, "word/header1.xml", { mode: "text", text: o.templateWatermark, seed: 1 }).xml;
    parts["word/header1.xml"] = header;
    parts["word/_rels/header1.xml.rels"] = rels([["rId1", "image", "media/logo.png"]]);
    parts["word/media/logo.png"] = "TPL-LOGO";
  }
  return parts;
}

/** Two sections: A4 portrait, then A4 landscape (final sectPr). */
function sourceParts(): Record<string, string> {
  const body =
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>`
    + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>`
    + `<w:p><w:r><w:t>Second section.</w:t></w:r></w:p>`
    + `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>`;
  return {
    "[Content_Types].xml": `<?xml version="1.0"?><Types ${CT}><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${WML}.header+xml"/><Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/></Types>`,
    "_rels/.rels": rels([["rId1", "officeDocument", "word/document.xml"]]),
    "word/document.xml": `<?xml version="1.0"?><w:document ${W} ${R}><w:body>${body}</w:body></w:document>`,
    "word/_rels/document.xml.rels": rels([["rId1", "styles", "styles.xml"], ["rId7", "header", "header1.xml"]]),
    "word/styles.xml": `<w:styles ${W}>${style("Normal", "Normal")}</w:styles>`,
    "word/header1.xml": `<w:hdr ${W}><w:p><w:r><w:t>SOURCE HEADER</w:t></w:r></w:p></w:hdr>`,
  };
}

const PORTRAIT_A4 = { w: 11906, h: 16838 };
const LANDSCAPE_A4 = { w: 16838, h: 11906 };
const PORTRAIT_LETTER = { w: 12240, h: 15840 };
const LANDSCAPE_LETTER = { w: 15840, h: 12240 };

const BASE_PAGE: PageSetupOptions = { preset: "template", orientation: "template", margins: "template", sections: "all" };

function formatWith(pageSetup: Partial<PageSetupOptions> = {}, watermark: WatermarkOptions = { mode: "inherit" }): FormatOptions {
  return { ...createDefaultFormatOptions(), pageSetup: { ...BASE_PAGE, ...pageSetup }, watermark };
}

async function text(docx: OpenedDocx, name: string): Promise<string> {
  const entry = docx.entries.get(name);
  if (entry === undefined) throw new Error("missing part " + name);
  return readEntryText(entry, 1 << 26);
}

interface Size { w: number; h: number; landscape: boolean; }

/** Every w:pgSz in document order. */
function pgSzs(main: string): Size[] {
  const out: Size[] = [];
  for (const m of main.matchAll(/<w:pgSz\b[^>]*>/g)) {
    const tag = m[0];
    const w = /\sw:w="(\d+)"/.exec(tag);
    const h = /\sw:h="(\d+)"/.exec(tag);
    out.push({ w: Number(w?.[1] ?? "0"), h: Number(h?.[1] ?? "0"), landscape: /w:orient="landscape"/.test(tag) });
  }
  return out;
}

/** Same checks the app runs on an output: validatePackage over the finished zip. */
async function validateOutput(blob: Blob): Promise<ValidationProblem[]> {
  const docx = await openDocx(blob, "out.docx");
  try {
    const partNames = [...docx.entries.keys()];
    const relsXml = new Map<string, string>();
    const facts = new Map<string, PartFacts>();
    for (const name of partNames) {
      if (name.endsWith(".rels")) relsXml.set(name, await text(docx, name));
      else if (/^word\/.*\.xml$/.test(name)) facts.set(name, scanXmlString(name, await text(docx, name)));
    }
    const numberingXml = docx.entries.has("word/numbering.xml") ? await text(docx, "word/numbering.xml") : undefined;
    const numberingIds = numberingXml === undefined
      ? undefined
      : new Set([...numberingXml.matchAll(/<w:num\s[^>]*w:numId="(\d+)"/g)].map((m) => Number(m[1] ?? "0")));
    const styleIds = facts.has("word/styles.xml") ? new Set(facts.get("word/styles.xml")?.styleDefIds ?? []) : undefined;
    return validatePackage({
      partNames,
      relsXml,
      contentTypesXml: docx.entries.has("[Content_Types].xml") ? await text(docx, "[Content_Types].xml") : undefined,
      facts,
      styleIds,
      numberingIds,
    });
  } finally {
    await docx.close();
  }
}

async function run(format: FormatOptions, tpl: TemplateOptions = {}, extra: { watermarkImage?: { data: Blob; extension: string } } = {}) {
  const options = extra.watermarkImage === undefined ? { format } : { format, watermarkImage: extra.watermarkImage };
  return applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts(tpl)), "tpl.docx", options);
}

async function partOf(blob: Blob, name: string): Promise<string> {
  const docx = await openDocx(blob, "out.docx");
  try {
    return await text(docx, name);
  } finally {
    await docx.close();
  }
}

async function codeOf(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    return isFormatIQError(error) ? error.code : "OTHER";
  }
  return "NO_ERROR";
}

describe("applyTemplate page setup", () => {
  it("default options keep the source sizes (A4 portrait, A4 landscape)", async () => {
    const result = await run(createDefaultFormatOptions());
    const sizes = pgSzs(await partOf(result.blob, "word/document.xml"));
    expect(sizes).toEqual([{ ...PORTRAIT_A4, landscape: false }, { ...LANDSCAPE_A4, landscape: true }]);
  });

  it("A4 to Letter: both sections change, the landscape one stays landscape", async () => {
    const result = await run(formatWith({ preset: "letter" }));
    const sizes = pgSzs(await partOf(result.blob, "word/document.xml"));
    expect(sizes[0]).toEqual({ ...PORTRAIT_LETTER, landscape: false });
    expect(sizes[1]).toEqual({ ...LANDSCAPE_LETTER, landscape: true });
    expect(result.report.pageSetupSectionsChanged).toBe(2);
  });

  it("orientation landscape turns every section landscape", async () => {
    const result = await run(formatWith({ orientation: "landscape" }));
    const sizes = pgSzs(await partOf(result.blob, "word/document.xml"));
    expect(sizes).toEqual([{ ...LANDSCAPE_A4, landscape: true }, { ...LANDSCAPE_A4, landscape: true }]);
  });

  it("orientation portrait turns every section portrait", async () => {
    const result = await run(formatWith({ orientation: "portrait" }));
    const sizes = pgSzs(await partOf(result.blob, "word/document.xml"));
    expect(sizes).toEqual([{ ...PORTRAIT_A4, landscape: false }, { ...PORTRAIT_A4, landscape: false }]);
  });

  it("writes margins into every section", async () => {
    const margins = { top: 1440, right: 1080, bottom: 1300, left: 1000, header: 700, footer: 650, gutter: 0 };
    const result = await run(formatWith({ margins }));
    const main = await partOf(result.blob, "word/document.xml");
    const tags = main.match(/<w:pgMar\b[^>]*>/g) ?? [];
    expect(tags.length).toBe(2);
    for (const tag of tags) {
      expect(tag).toContain('w:top="1440"');
      expect(tag).toContain('w:right="1080"');
      expect(tag).toContain('w:bottom="1300"');
      expect(tag).toContain('w:left="1000"');
      expect(tag).toContain('w:header="700"');
      expect(tag).toContain('w:footer="650"');
    }
    expect(result.report.pageSetupSectionsChanged).toBe(2);
  });

  it("selected sections: only section 0 changes", async () => {
    const result = await run(formatWith({ preset: "letter", sections: [0] }));
    const sizes = pgSzs(await partOf(result.blob, "word/document.xml"));
    expect(sizes[0]).toEqual({ ...PORTRAIT_LETTER, landscape: false });
    expect(sizes[1]).toEqual({ ...LANDSCAPE_A4, landscape: true });
    expect(result.report.pageSetupSectionsChanged).toBe(1);
  });

  it("custom size is used", async () => {
    const result = await run(formatWith({ preset: "custom", customWidthTwips: 10000, customHeightTwips: 14000 }));
    const sizes = pgSzs(await partOf(result.blob, "word/document.xml"));
    expect(sizes[0]).toEqual({ w: 10000, h: 14000, landscape: false });
  });

  it("the output with page setup passes validatePackage with no errors", async () => {
    const result = await run(formatWith({ preset: "letter", orientation: "landscape" }));
    const problems = await validateOutput(result.blob);
    expect(problems.filter((p) => p.severity === "error")).toEqual([]);
  });
});

describe("applyTemplate watermark", () => {
  it("inherit leaves the template watermark alone", async () => {
    const result = await run(formatWith({}, { mode: "inherit" }), { templateWatermark: "TEMPLATE-WM" });
    expect(await partOf(result.blob, "word/header1.xml")).toContain("TEMPLATE-WM");
    expect(result.report.watermarksAdded).toBe(0);
    expect(result.report.watermarksRemoved).toBe(0);
  });

  it("remove takes the template watermark out", async () => {
    const result = await run(formatWith({}, { mode: "remove" }), { templateWatermark: "TEMPLATE-WM" });
    const header = await partOf(result.blob, "word/header1.xml");
    expect(header).not.toContain("TEMPLATE-WM");
    expect(header).toContain("TEMPLATE HEADER");
    expect(result.report.watermarksRemoved).toBe(1);
    expect(result.report.watermarksAdded).toBe(0);
  });

  it("text replaces the template watermark with a new one", async () => {
    const result = await run(formatWith({}, { mode: "text", text: "NEW-WM" }), { templateWatermark: "TEMPLATE-WM" });
    const header = await partOf(result.blob, "word/header1.xml");
    expect(header).toContain("NEW-WM");
    expect(header).not.toContain("TEMPLATE-WM");
    expect(result.report.watermarksRemoved).toBe(1);
    expect(result.report.watermarksAdded).toBe(1);
  });

  it("text on a header without a watermark adds one", async () => {
    const result = await run(formatWith({}, { mode: "text", text: "DRAFT" }));
    expect(await partOf(result.blob, "word/header1.xml")).toContain("DRAFT");
    expect(result.report.watermarksAdded).toBe(1);
    expect(result.report.watermarksRemoved).toBe(0);
    expect((await validateOutput(result.blob)).filter((p) => p.severity === "error")).toEqual([]);
  });

  it("keepSource warns and keeps the template's headers and watermark", async () => {
    const result = await run(formatWith({}, { mode: "keepSource" }), { templateWatermark: "TEMPLATE-WM" });
    expect(await partOf(result.blob, "word/header1.xml")).toContain("TEMPLATE-WM");
    expect(result.report.warnings.join(" ")).toContain("source watermark");
  });

  it("a template with no header plus text gives the warning and adds nothing", async () => {
    const result = await run(formatWith({}, { mode: "text", text: "DRAFT" }), { header: false });
    expect(result.report.watermarksAdded).toBe(0);
    expect(result.report.warnings.join(" ")).toContain("no header");
    const docx = await openDocx(result.blob, "out.docx");
    expect([...docx.entries.keys()].some((n) => /^word\/header/.test(n))).toBe(false);
    await docx.close();
  });

  it("image mode adds the relationship, the picture and the content type exactly once", async () => {
    const image = { data: new Blob(["PNGDATA"]), extension: "png" };
    const result = await run(formatWith({}, { mode: "image", widthPt: 200, heightPt: 100 }), {}, { watermarkImage: image });
    expect(result.report.watermarksAdded).toBe(1);
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/media/formatiq_watermark.png")).toBe("PNGDATA");
    const headerRels = parseRelationships(await text(docx, "word/_rels/header1.xml.rels"));
    const pictureRels = headerRels.filter((r) => r.target.endsWith("media/formatiq_watermark.png"));
    expect(pictureRels.length).toBe(1);
    expect(headerRels.some((r) => r.target.endsWith("media/logo.png"))).toBe(true);
    const header = await text(docx, "word/header1.xml");
    expect(header).toContain(`"${pictureRels[0]?.id ?? "MISSING"}"`);
    const types = await text(docx, "[Content_Types].xml");
    expect(types.match(/Extension="png"/g)?.length).toBe(1);
    await docx.close();
    expect((await validateOutput(result.blob)).filter((p) => p.severity === "error")).toEqual([]);
  });

  it("image mode without a picture, or with an unknown picture type, is refused", async () => {
    const wm: WatermarkOptions = { mode: "image", widthPt: 200, heightPt: 100 };
    expect(await codeOf(run(formatWith({}, wm)))).toBe("INVALID_OPTIONS");
    expect(await codeOf(run(formatWith({}, wm), {}, { watermarkImage: { data: new Blob(["x"]), extension: "tiff" } }))).toBe("INVALID_OPTIONS");
  });
});
