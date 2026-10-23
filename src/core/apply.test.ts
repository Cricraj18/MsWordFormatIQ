import { BlobWriter, TextReader, ZipWriter } from "@zip.js/zip.js";
import { beforeAll, describe, expect, it } from "vitest";
import { isFormatIQError } from "../types";
import type { DetectionResult } from "../types";
import { applyTemplate } from "./apply";
import { detectStructure } from "./structure_detect";
import { readParagraphs } from "./paragraphs";
import { openDocx, readEntryText, setWorkersEnabled } from "./zipio";
import type { OpenedDocx } from "./zipio";
import { parseContentTypes, parseRelationships, contentTypeOf, resolveTarget } from "./package_merge";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
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

function rels(items: [string, string, string, boolean?][]): string {
  const body = items.map(([id, type, target, ext]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${ext === true ? ' TargetMode="External"' : ""}/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships ${PK}>${body}</Relationships>`;
}

function style(id: string, name: string): string {
  return `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/></w:style>`;
}

const STYLES = `<w:styles ${W}>${style("Normal", "Normal")}${style("Heading1", "heading 1")}${style("Title", "Title")}${style("ListParagraph", "List Paragraph")}${style("Caption", "caption")}</w:styles>`;

const TPL_SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:footerReference w:type="default" r:id="rId11"/><w:titlePg/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

interface TemplateOptions { numbering?: boolean; }

function templateParts(o: TemplateOptions = {}): Record<string, string> {
  const withNum = o.numbering !== false;
  const numRel: [string, string, string][] = withNum ? [["rId2", "numbering", "numbering.xml"]] : [];
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types ${CT}><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${WML}.settings+xml"/><Override PartName="/word/header1.xml" ContentType="${WML}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="${WML}.footer+xml"/>${withNum ? `<Override PartName="/word/numbering.xml" ContentType="${WML}.numbering+xml"/>` : ""}</Types>`,
    "_rels/.rels": rels([["rId1", "officeDocument", "word/document.xml"]]),
    "word/document.xml": `<?xml version="1.0"?><w:document ${W} ${R}><w:body><w:p><w:r><w:t>Template body text</w:t></w:r></w:p>${TPL_SECT}</w:body></w:document>`,
    "word/_rels/document.xml.rels": rels([["rId1", "styles", "styles.xml"], ...numRel, ["rId3", "settings", "settings.xml"], ["rId10", "header", "header1.xml"], ["rId11", "footer", "footer1.xml"], ["rId12", "image", "media/tplbody.png"]]),
    "word/styles.xml": STYLES,
    "word/settings.xml": `<w:settings ${W}><w:zoom w:percent="100"/><w:compat/></w:settings>`,
    "word/header1.xml": `<w:hdr ${W}><w:p><w:r><w:t>TEMPLATE HEADER</w:t></w:r></w:p></w:hdr>`,
    "word/_rels/header1.xml.rels": rels([["rId1", "image", "media/logo.png"]]),
    "word/footer1.xml": `<w:ftr ${W}><w:p><w:r><w:t>TEMPLATE FOOTER</w:t></w:r></w:p></w:ftr>`,
    "word/media/logo.png": "TPL-LOGO",
    "word/media/tplbody.png": "TPL-BODY-IMAGE",
  };
  if (withNum) {
    parts["word/numbering.xml"] = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  }
  return parts;
}

const SRC_BODY = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>`
  + `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="120"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Comic Sans MS"/><w:sz w:val="40"/></w:rPr><w:t>Plain body text here.</w:t></w:r></w:p>`
  + `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>`
  + `<w:p><w:r><w:drawing><a:blip r:embed="rId5"/></w:drawing></w:r></w:p>`
  + `<w:p><w:r><w:t>See note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:hyperlink r:id="rId6"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`
  + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>`
  + `<w:p><w:r><w:t>Second section text.</w:t></w:r></w:p>`
  + `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>`;

function sourceParts(body: string = SRC_BODY): Record<string, string> {
  return {
    "[Content_Types].xml": `<?xml version="1.0"?><Types ${CT}><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="${WML}.footnotes+xml"/><Override PartName="/word/header1.xml" ContentType="${WML}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="${WML}.footer+xml"/><Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${WML}.numbering+xml"/></Types>`,
    "_rels/.rels": rels([["rId1", "officeDocument", "word/document.xml"]]),
    "word/document.xml": `<?xml version="1.0"?><w:document ${W} ${R} ${A}><w:body>${body}</w:body></w:document>`,
    "word/_rels/document.xml.rels": rels([["rId1", "styles", "styles.xml"], ["rId2", "numbering", "numbering.xml"], ["rId5", "image", "media/logo.png"], ["rId6", "hyperlink", "https://example.com", true], ["rId7", "header", "header1.xml"], ["rId8", "footnotes", "footnotes.xml"], ["rId9", "footer", "footer1.xml"]]),
    "word/styles.xml": `<w:styles ${W}>${style("Normal", "Normal")}</w:styles>`,
    "word/numbering.xml": `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    "word/header1.xml": `<w:hdr ${W}><w:p><w:r><w:t>SOURCE HEADER</w:t></w:r></w:p></w:hdr>`,
    "word/footer1.xml": `<w:ftr ${W}><w:p><w:r><w:t>SOURCE FOOTER</w:t></w:r></w:p></w:ftr>`,
    "word/footnotes.xml": `<w:footnotes ${W}><w:footnote w:id="1"><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Note text</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    "word/_rels/footnotes.xml.rels": rels([["rId1", "image", "media/fn.png"]]),
    "word/media/logo.png": "SRC-LOGO",
    "word/media/fn.png": "FN-IMAGE",
  };
}

async function text(docx: OpenedDocx, name: string): Promise<string> {
  const entry = docx.entries.get(name);
  if (entry === undefined) throw new Error("missing part " + name);
  return readEntryText(entry, 1 << 26);
}

async function paragraphTexts(blob: Blob): Promise<string[]> {
  const docx = await openDocx(blob, "x.docx");
  const out: string[] = [];
  await readParagraphs(docx, (p) => { out.push(p.text); });
  await docx.close();
  return out;
}

/** Structural checks: rels targets exist, every r:id of the main part resolves, every part has a content type. */
async function checkPackage(blob: Blob): Promise<string[]> {
  const problems: string[] = [];
  const docx = await openDocx(blob, "out.docx");
  const ct = parseContentTypes(await text(docx, "[Content_Types].xml"));
  for (const name of docx.entries.keys()) {
    if (name === "[Content_Types].xml") continue;
    if (contentTypeOf(ct, name) === undefined) problems.push("no content type: " + name);
    if (name.endsWith(".rels")) {
      const owner = name.replace(/_rels\//, "").replace(/\.rels$/, "");
      const dir = owner.includes("/") ? owner.slice(0, owner.lastIndexOf("/")) : "";
      const items = name === "_rels/.rels" ? parseRelationships(await text(docx, name)) : parseRelationships(await text(docx, name));
      for (const r of items) {
        if (r.targetMode === "External") continue;
        const target = resolveTarget(name === "_rels/.rels" ? "" : dir, r.target);
        if (target === undefined || !docx.entries.has(target)) problems.push(`${name}: missing target ${r.target}`);
      }
    }
  }
  for (const key of ct.overrides.keys()) if (!docx.entries.has(key.slice(1))) problems.push("override for missing part " + key);
  const main = await text(docx, "word/document.xml");
  const ids = new Set(parseRelationships(await text(docx, "word/_rels/document.xml.rels")).map((r) => r.id));
  for (const m of main.matchAll(/r:(?:id|embed)="([^"]+)"/g)) if (!ids.has(m[1] ?? "")) problems.push("unresolved r:id " + (m[1] ?? ""));
  await docx.close();
  return problems;
}

describe("applyTemplate", () => {
  it("builds a package that passes the structural checks", async () => {
    const source = await makeZip(sourceParts());
    const result = await applyTemplate(source, "src.docx", await makeZip(templateParts()), "tpl.docx");
    expect(await checkPackage(result.blob)).toEqual([]);
  });

  it("restyles paragraphs from detected roles and clears direct formatting", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    const docx = await openDocx(result.blob, "out.docx");
    const main = await text(docx, "word/document.xml");
    await docx.close();
    expect(main).toContain('<w:pStyle w:val="Heading1"/>');
    expect(main).toContain('<w:pStyle w:val="Normal"/>');
    expect(main).toContain('<w:pStyle w:val="ListParagraph"/>');
    expect(main).not.toContain("w:jc");
    expect(main).not.toContain("Comic Sans");
    expect(result.report.paragraphs).toBe(7);
    expect(result.report.clearedParagraphProps).toBeGreaterThanOrEqual(2);
    expect(result.report.clearedRunProps).toBeGreaterThanOrEqual(2);
  });

  it("keeps the text of every paragraph and detects the same roles again", async () => {
    const source = await makeZip(sourceParts());
    const result = await applyTemplate(source, "src.docx", await makeZip(templateParts()), "tpl.docx");
    expect(await paragraphTexts(result.blob)).toEqual(await paragraphTexts(source));
    const again = await detectStructure(result.blob, "out.docx");
    const roles = again.paragraphs.map((p) => p.role);
    expect(roles[0]).toBe("heading");
    expect(roles[1]).toBe("body");
    expect(roles[2]).toBe("list-item");
  });

  it("uses the template's styles.xml, settings and header/footer parts, not the source's", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/styles.xml")).toBe(STYLES);
    expect(await text(docx, "word/header1.xml")).toContain("TEMPLATE HEADER");
    expect(await text(docx, "word/footer1.xml")).toContain("TEMPLATE FOOTER");
    const all = [...docx.entries.keys()];
    expect(all).not.toContain("word/media/tplbody.png");
    expect(all).toContain("word/media/logo.png");
    await docx.close();
  });

  it("replaces source header/footer references by the template's, in every section", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    const docx = await openDocx(result.blob, "out.docx");
    const main = await text(docx, "word/document.xml");
    await docx.close();
    expect(main).not.toContain('r:id="rId7"');
    expect(main).not.toContain('r:id="rId9"');
    expect(main.match(/<w:headerReference w:type="default" r:id="tplrId10"\/>/g)?.length).toBe(2);
    expect(main.match(/<w:footerReference w:type="default" r:id="tplrId11"\/>/g)?.length).toBe(2);
    expect(main.match(/<w:titlePg\/>/g)?.length).toBe(2);
    expect(result.report.sectionsFixed).toBe(2);
    expect(result.report.sourceHeaderFooterRefsRemoved).toBe(3);
  });

  it("keeps the source page size of each section (landscape stays)", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    const docx = await openDocx(result.blob, "out.docx");
    const main = await text(docx, "word/document.xml");
    await docx.close();
    expect(main).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
  });

  it("renames a source part that clashes with a template part and fixes the link", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    expect(result.report.renamedParts).toEqual([["word/media/logo.png", "word/media/src_logo.png"]]);
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/media/logo.png")).toBe("TPL-LOGO");
    expect(await text(docx, "word/media/src_logo.png")).toBe("SRC-LOGO");
    const docRels = parseRelationships(await text(docx, "word/_rels/document.xml.rels"));
    expect(docRels.find((r) => r.id === "rId5")?.target).toBe("media/src_logo.png");
    expect(docRels.find((r) => r.id === "rId6")?.targetMode).toBe("External");
    expect(docRels.find((r) => r.id === "rId8")?.type.endsWith("/footnotes")).toBe(true);
    expect(docRels.find((r) => r.id === "tplrId1")?.type.endsWith("/styles")).toBe(true);
    expect(docRels.some((r) => r.id === "rId1" || r.id === "rId7" || r.id === "rId9")).toBe(false);
    await docx.close();
  });

  it("follows relationships of copied parts (image used only by footnotes)", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/media/fn.png")).toBe("FN-IMAGE");
    const fnRels = parseRelationships(await text(docx, "word/_rels/footnotes.xml.rels"));
    expect(fnRels[0]?.target).toBe("media/fn.png");
    await docx.close();
  });

  it("merges numbering and remaps list ids in the document and in footnotes", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    const docx = await openDocx(result.blob, "out.docx");
    const main = await text(docx, "word/document.xml");
    const notes = await text(docx, "word/footnotes.xml");
    const numbering = await text(docx, "word/numbering.xml");
    await docx.close();
    expect(main).toContain('<w:numId w:val="2"/>');
    expect(notes).toContain('<w:numId w:val="2"/>');
    expect(numbering.match(/<w:num /g)?.length).toBe(2);
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>');
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>');
    expect(result.report.numIdsRemapped).toBe(2);
    expect(result.report.numIdsUnmapped).toBe(0);
  });

  it("adds a numbering part when only the source has lists", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts({ numbering: false })), "tpl.docx");
    expect(await checkPackage(result.blob)).toEqual([]);
    const docx = await openDocx(result.blob, "out.docx");
    expect(docx.entries.has("word/numbering.xml")).toBe(true);
    const rel = parseRelationships(await text(docx, "word/_rels/document.xml.rels")).find((r) => r.type.endsWith("/numbering"));
    expect(rel?.target).toBe("numbering.xml");
    expect(parseContentTypes(await text(docx, "[Content_Types].xml")).overrides.get("/word/numbering.xml")).toContain("numbering+xml");
    expect(await text(docx, "word/document.xml")).toContain('<w:numId w:val="1"/>');
    await docx.close();
  });

  it("sets update-fields-on-open only when the document has fields", async () => {
    const plain = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    expect(plain.report.fieldsUpdateOnOpen).toBe(false);
    const withField = SRC_BODY.replace("Second section text.", "Page <w:fldSimple w:instr=\" PAGE \"><w:r><w:t>1</w:t></w:r></w:fldSimple>");
    const result = await applyTemplate(await makeZip(sourceParts(withField)), "src.docx", await makeZip(templateParts()), "tpl.docx");
    expect(result.report.fieldsUpdateOnOpen).toBe(true);
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/settings.xml")).toBe(`<w:settings ${W}><w:zoom w:percent="100"/><w:updateFields w:val="true"/><w:compat/></w:settings>`);
    await docx.close();
  });

  it("uses a given detection with manual tags", async () => {
    const source = await makeZip(sourceParts());
    const detection: DetectionResult = await detectStructure(source, "src.docx");
    const first = detection.paragraphs[1];
    if (first === undefined) throw new Error("no paragraph");
    first.manualRole = "title";
    const result = await applyTemplate(source, "src.docx", await makeZip(templateParts()), "tpl.docx", { detection });
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/document.xml")).toContain('<w:pStyle w:val="Title"/>');
    await docx.close();
  });

  it("keeps direct formatting when asked", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx", { directFormatting: "keep" });
    const docx = await openDocx(result.blob, "out.docx");
    expect(await text(docx, "word/document.xml")).toContain("Comic Sans");
    await docx.close();
  });

  it("does not change the source blob and gives the same result twice", async () => {
    const source = await makeZip(sourceParts());
    const before = await source.arrayBuffer();
    const template = await makeZip(templateParts());
    const one = await applyTemplate(source, "src.docx", template, "tpl.docx");
    const two = await applyTemplate(source, "src.docx", template, "tpl.docx");
    expect(Buffer.from(await source.arrayBuffer()).equals(Buffer.from(before))).toBe(true);
    expect(one.report).toEqual(two.report);
    const a = await openDocx(one.blob, "a.docx");
    const b = await openDocx(two.blob, "b.docx");
    expect(await text(a, "word/document.xml")).toBe(await text(b, "word/document.xml"));
    await a.close();
    await b.close();
  });

  it("reports template styles that are missing", async () => {
    const tpl = templateParts();
    tpl["word/styles.xml"] = `<w:styles ${W}>${style("Normal", "Normal")}</w:styles>`;
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(tpl), "tpl.docx");
    expect(result.report.missingStyles).toEqual(["heading 1", "list paragraph"]);
  });

  it("reports the things that are not carried over", async () => {
    const result = await applyTemplate(await makeZip(sourceParts()), "src.docx", await makeZip(templateParts()), "tpl.docx");
    expect(result.report.droppedSourceRelationships).toBe(4);
    expect(result.report.warnings.join(" ")).toContain("not carried over");
  });
});

describe("applyTemplate errors", () => {
  async function fails(promise: Promise<unknown>): Promise<string> {
    try {
      await promise;
    } catch (error) {
      return isFormatIQError(error) ? error.code : "OTHER";
    }
    return "NO_ERROR";
  }

  it("refuses scope selected for now", async () => {
    const options = { format: { scope: "selected" as const, parts: ["styles" as const], includeToc: false, includeListOfFigures: false, includeListOfTables: false, updateFieldsOnOpen: false, imageReduction: "original" as const } };
    expect(await fails(applyTemplate(await makeZip(sourceParts()), "s.docx", await makeZip(templateParts()), "t.docx", options))).toBe("INVALID_OPTIONS");
  });

  it("refuses a source that is not .docx", async () => {
    expect(await fails(applyTemplate(await makeZip(sourceParts()), "s.pdf", await makeZip(templateParts()), "t.docx"))).toBe("NOT_DOCX");
  });

  it("refuses a template without styles.xml", async () => {
    const tpl = templateParts();
    delete tpl["word/styles.xml"];
    expect(await fails(applyTemplate(await makeZip(sourceParts()), "s.docx", await makeZip(tpl), "t.docx"))).toBe("TEMPLATE_INVALID");
  });

  it("refuses a source with damaged XML", async () => {
    const parts = sourceParts();
    parts["word/document.xml"] = "<w:document><w:body><w:p></w:body>";
    expect(await fails(applyTemplate(await makeZip(parts), "s.docx", await makeZip(templateParts()), "t.docx"))).toBe("ZIP_READ_FAILED");
  });

  it("stops when cancelled", async () => {
    const controller = new AbortController();
    controller.abort();
    expect(await fails(applyTemplate(await makeZip(sourceParts()), "s.docx", await makeZip(templateParts()), "t.docx", { signal: controller.signal }))).toBe("CANCELLED");
  });
});
