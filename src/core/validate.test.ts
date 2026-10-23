import { describe, expect, it } from "vitest";
import { BlobReader, TextWriter, ZipReader, configure } from "@zip.js/zip.js";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { listNumIds } from "./numbering_merge";
import {
  PartFacts,
  ValidationInput,
  createPartScanner,
  hasErrors,
  scanXmlString,
  validatePackage,
} from "./validate";

configure({ useWebWorkers: false });

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_XML =
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="png" ContentType="image/png"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';

function rels(...items: [string, string, string][]): string {
  return `<Relationships xmlns="${RELS_NS}">` + items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("") + "</Relationships>";
}

function body(inner: string): string {
  return `<w:document ${W}><w:body>${inner}</w:body></w:document>`;
}

function base(overrides: Partial<ValidationInput> = {}, docXml = body("<w:p/>")): ValidationInput {
  return {
    partNames: ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/_rels/document.xml.rels", "word/media/a.png"],
    relsXml: new Map([
      ["_rels/.rels", rels(["rId1", "x/officeDocument", "word/document.xml"])],
      ["word/_rels/document.xml.rels", rels(["rId5", "x/image", "media/a.png"])],
    ]),
    contentTypesXml: CT_XML,
    facts: new Map<string, PartFacts>([["word/document.xml", scanXmlString("word/document.xml", docXml)]]),
    styleIds: new Set(["Normal", "Heading1"]),
    numberingIds: new Set([1, 2]),
    ...overrides,
  };
}

const codes = (input: ValidationInput): string[] => validatePackage(input).map((p) => p.code);

describe("scanXmlString", () => {
  it("collects ids, styles, lists, bookmarks and sections", () => {
    const f = scanXmlString("d", body(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>' +
      '<w:bookmarkStart w:id="1" w:name="a"/><w:r><w:drawing><a:blip xmlns:a="urn:a" r:embed="rId5"/></w:drawing></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'));
    expect([...f.rIds].sort()).toEqual(["rId5", "rId9"]);
    expect([...f.pStyleIds]).toEqual(["Heading1"]);
    expect([...f.numIds]).toEqual([3]);
    expect(f.bookmarkIds).toEqual(["1"]);
    expect(f.bookmarkNames).toEqual(["a"]);
    expect(f.sectPrCount).toBe(1);
    expect(f.sectPrIssues).toEqual([]);
  });
  it("ignores numId 0 and attributes of other namespaces", () => {
    const f = scanXmlString("d", `<w:p ${W} xmlns:x="urn:x"><w:numId w:val="0"/><w:foo x:id="zzz"/></w:p>`);
    expect(f.numIds.size).toBe(0);
    expect(f.rIds.size).toBe(0);
  });
  it("reads style definitions", () => {
    expect([...scanXmlString("s", `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Normal"/></w:styles>`).styleDefIds]).toEqual(["Normal"]);
  });
  it("works when fed in small pieces", () => {
    const xml = body('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>');
    const s = createPartScanner("d");
    for (let i = 0; i < xml.length; i += 7) s.write(xml.slice(i, i + 7));
    expect([...s.end().pStyleIds]).toEqual(["Heading1"]);
  });
  it("flags sections without or with bad page size and repeated references", () => {
    expect(scanXmlString("d", body("<w:sectPr/>")).sectPrIssues.join(" ")).toMatch(/no page size/);
    expect(scanXmlString("d", body('<w:sectPr><w:pgSz w:w="0" w:h="5"/></w:sectPr>')).sectPrIssues.join(" ")).toMatch(/not a positive/);
    const twice = '<w:sectPr><w:headerReference w:type="first" r:id="a"/><w:headerReference w:type="first" r:id="b"/><w:pgSz w:w="1" w:h="1"/></w:sectPr>';
    expect(scanXmlString("d", body(twice)).sectPrIssues.join(" ")).toMatch(/2 header\/footer/);
  });
  it("handles a section inside a section change", () => {
    const x = body('<w:sectPr><w:pgSz w:w="1" w:h="1"/><w:sectPrChange w:id="1"><w:sectPr><w:pgSz w:w="2" w:h="2"/></w:sectPr></w:sectPrChange></w:sectPr>');
    const f = scanXmlString("d", x);
    expect(f.sectPrCount).toBe(2);
    expect(f.sectPrIssues).toEqual([]);
  });
  it("throws for damaged XML", () => {
    expect(() => scanXmlString("d", "<w:p><w:r></w:p>")).toThrow();
  });
});

describe("validatePackage", () => {
  it("accepts a clean package", () => {
    expect(validatePackage(base())).toEqual([]);
    expect(hasErrors([])).toBe(false);
  });
  it("finds missing main part, root rels and content types", () => {
    const input = base({ partNames: ["word/other.xml"], relsXml: new Map(), contentTypesXml: undefined, facts: new Map() });
    expect(codes(input)).toEqual(expect.arrayContaining(["MISSING_MAIN", "MISSING_ROOT_RELS", "MISSING_CONTENT_TYPES"]));
  });
  it("finds part names that differ only by case", () => {
    expect(codes(base({ partNames: ["word/document.xml", "word/Document.xml", "_rels/.rels", "[Content_Types].xml"] }))).toContain("DUPLICATE_PART");
  });
  it("finds a part without content type and an override for a missing part", () => {
    const ct = CT_XML.replace('<Default Extension="png" ContentType="image/png"/>', "").replace("</Types>", '<Override PartName="/word/gone.xml" ContentType="x"/></Types>');
    const c = codes(base({ contentTypesXml: ct }));
    expect(c).toContain("NO_CONTENT_TYPE");
    expect(c).toContain("OVERRIDE_WITHOUT_PART");
    expect(validatePackage(base({ contentTypesXml: ct })).find((p) => p.code === "OVERRIDE_WITHOUT_PART")?.severity).toBe("warning");
  });
  it("finds relationship targets that are missing or outside", () => {
    const input = base({
      relsXml: new Map([
        ["_rels/.rels", rels(["rId1", "x/officeDocument", "word/document.xml"])],
        ["word/_rels/document.xml.rels", rels(["rId5", "x/image", "media/missing.png"], ["rId6", "x/image", "../../../x.png"])],
      ]),
    });
    expect(codes(input)).toEqual(expect.arrayContaining(["MISSING_TARGET", "REL_TARGET_OUTSIDE"]));
  });
  it("ignores external targets and finds duplicate relationship ids", () => {
    const ext = `<Relationships xmlns="${RELS_NS}"><Relationship Id="rId5" Type="x/hyperlink" Target="http://a.b" TargetMode="External"/></Relationships>`;
    expect(codes(base({ relsXml: new Map([["_rels/.rels", rels(["rId1", "x/officeDocument", "word/document.xml"])], ["word/_rels/document.xml.rels", ext]]) }))).toEqual([]);
    const dup = rels(["rId5", "x/image", "media/a.png"], ["rId5", "x/image", "media/a.png"]);
    expect(codes(base({ relsXml: new Map([["_rels/.rels", rels(["rId1", "x/officeDocument", "word/document.xml"])], ["word/_rels/document.xml.rels", dup]]) }))).toContain("DUPLICATE_REL_ID");
  });
  it("finds unresolved r:id, also when the part has no rels file", () => {
    const doc = body('<w:p><w:r><w:drawing><a:blip xmlns:a="urn:a" r:embed="rId99"/></w:drawing></w:r></w:p>');
    expect(codes(base({}, doc))).toContain("UNRESOLVED_RID");
    const noRels = base({ relsXml: new Map([["_rels/.rels", rels(["rId1", "x/officeDocument", "word/document.xml"])]]) }, doc);
    expect(validatePackage(noRels).find((p) => p.code === "UNRESOLVED_RID")?.message).toMatch(/no relationships file/);
  });
  it("resolves r:id in headers with their own rels", () => {
    const header = `<w:hdr ${W}><a:blip xmlns:a="urn:a" r:embed="rId1"/></w:hdr>`;
    const input = base({
      partNames: ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/header1.xml", "word/_rels/header1.xml.rels", "word/media/a.png"],
      relsXml: new Map([
        ["_rels/.rels", rels(["rId1", "x/officeDocument", "word/document.xml"])],
        ["word/_rels/header1.xml.rels", rels(["rId1", "x/image", "media/a.png"])],
      ]),
      facts: new Map([["word/document.xml", scanXmlString("d", body("<w:p/>"))], ["word/header1.xml", scanXmlString("h", header)]]),
    });
    expect(codes(input)).toEqual([]);
  });
  it("warns about unknown paragraph styles and errors on unknown list ids", () => {
    const doc = body('<w:p><w:pPr><w:pStyle w:val="Nope"/><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:p>');
    const list = validatePackage(base({}, doc));
    expect(list.find((p) => p.code === "UNKNOWN_STYLE")?.severity).toBe("warning");
    expect(list.find((p) => p.code === "UNKNOWN_NUMID")?.severity).toBe("error");
    expect(hasErrors(list)).toBe(true);
  });
  it("reports list ids when there is no numbering part, and skips style check without styles", () => {
    const doc = body('<w:p><w:pPr><w:pStyle w:val="Nope"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>');
    const c = codes(base({ numberingIds: undefined, styleIds: undefined }, doc));
    expect(c).toContain("UNKNOWN_NUMID");
    expect(c).not.toContain("UNKNOWN_STYLE");
  });
  it("finds duplicate bookmark ids (error) and names (warning)", () => {
    const doc = body('<w:bookmarkStart w:id="1" w:name="a"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="1" w:name="a"/><w:bookmarkEnd w:id="1"/>');
    const list = validatePackage(base({}, doc));
    expect(list.find((p) => p.code === "DUPLICATE_BOOKMARK_ID")?.severity).toBe("error");
    expect(list.find((p) => p.code === "DUPLICATE_BOOKMARK_NAME")?.severity).toBe("warning");
  });
  it("reports section problems as warnings", () => {
    const list = validatePackage(base({}, body("<w:sectPr/>")));
    expect(list.map((p) => p.code)).toEqual(["BAD_SECTPR"]);
    expect(hasErrors(list)).toBe(false);
  });
  it("accepts numbering ids from listNumIds", () => {
    const numbering = `<w:numbering ${W}><w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const doc = body('<w:p><w:pPr><w:numPr><w:numId w:val="4"/></w:numPr></w:pPr></w:p>');
    expect(codes(base({ numberingIds: listNumIds(numbering) }, doc))).toEqual([]);
  });
});

describe("validatePackage on a real generated .docx", () => {
  it("finds no errors in a file made by the docx library", async () => {
    const d = new Document({ sections: [{ children: [new Paragraph({ children: [new TextRun("Hello")] })] }] });
    const blob = new Blob([new Uint8Array(await Packer.toBuffer(d))]);
    const reader = new ZipReader(new BlobReader(blob));
    const entries = await reader.getEntries();
    const partNames: string[] = [];
    const relsXml = new Map<string, string>();
    const facts = new Map<string, PartFacts>();
    let contentTypesXml: string | undefined;
    let styleIds: Set<string> | undefined;
    let numberingIds: Set<number> | undefined;
    for (const e of entries) {
      partNames.push(e.filename);
      if (e.directory || e.getData === undefined) continue;
      const text = await e.getData(new TextWriter());
      if (e.filename === "[Content_Types].xml") contentTypesXml = text;
      else if (e.filename.endsWith(".rels")) relsXml.set(e.filename, text);
      else if (e.filename.endsWith(".xml")) {
        const f = scanXmlString(e.filename, text);
        facts.set(e.filename, f);
        if (e.filename === "word/styles.xml") styleIds = f.styleDefIds;
        if (e.filename === "word/numbering.xml") numberingIds = listNumIds(text);
      }
    }
    await reader.close();
    const list = validatePackage({ partNames, relsXml, contentTypesXml, facts, styleIds, numberingIds });
    expect(list.filter((p) => p.severity === "error")).toEqual([]);
    expect(partNames).toContain("word/document.xml");
  });
});
