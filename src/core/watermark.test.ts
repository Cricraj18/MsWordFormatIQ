import { describe, expect, it } from "vitest";
import { FormatIQError } from "../types";
import {
  IMAGE_MARKER,
  TEXT_MARKER,
  addImageRelationship,
  buildImageWatermarkRun,
  buildTextWatermarkRun,
  createRunInserter,
  findWatermarks,
  formatOpacity,
  removeWatermarks,
  setWatermark,
} from "./watermark";
import type { WatermarkOptions } from "./watermark";
import { parseRelationships } from "./package_merge";
import { createXmlRewriter, rewriteXmlString } from "./xml_rewrite";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HDR_OPEN = `<w:hdr xmlns:w="${W}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;
const BARE_OPEN = `<w:hdr xmlns:w="${W}">`;
const para = (inner: string, ppr = `<w:pPr><w:pStyle w:val="Header"/></w:pPr>`): string => `<w:p>${ppr}${inner}</w:p>`;
const textRun = (t: string): string => `<w:r><w:t>${t}</w:t></w:r>`;

const OLD_TEXT_WM =
  `<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600"><v:path textpathok="t"/></v:shapetype>` +
  `<v:shape id="PowerPlusWaterMarkObject357476642" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;width:412pt;height:247pt;rotation:315;z-index:-251658752" fillcolor="silver" stroked="f">` +
  `<v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r>`;
const OLD_IMAGE_WM =
  `<w:r><w:pict><v:shape id="WordPictureWatermark1" type="#_x0000_t75" style="position:absolute;width:300pt;height:200pt"><v:imagedata r:id="rId7" o:title=""/></v:shape></w:pict></w:r>`;

const opts = (over: Partial<WatermarkOptions>): WatermarkOptions => ({ mode: "text", text: "CONFIDENTIAL", ...over });

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    return error instanceof FormatIQError ? error.code : "OTHER";
  }
  return "NONE";
}

describe("finding watermarks", () => {
  it("reads a text watermark", () => {
    const list = findWatermarks(`${HDR_OPEN}${para(OLD_TEXT_WM)}</w:hdr>`);
    expect(list).toEqual([
      { kind: "text", id: "PowerPlusWaterMarkObject357476642", text: "DRAFT", color: "silver", opacity: ".5", rotated: true, relId: undefined },
    ]);
  });
  it("reads an image watermark and its relationship id", () => {
    const list = findWatermarks(`${HDR_OPEN}${para(OLD_IMAGE_WM)}</w:hdr>`);
    expect(list).toHaveLength(1);
    expect(list[0]).toMatchObject({ kind: "image", relId: "rId7", rotated: false, text: undefined });
  });
  it("finds nothing in a plain header or for a shape with another id", () => {
    expect(findWatermarks(`${HDR_OPEN}${para(textRun("Page"))}</w:hdr>`)).toEqual([]);
    expect(findWatermarks(`${HDR_OPEN}${para(`<w:r><w:pict><v:shape id="Logo1"><v:textpath string="x"/></v:shape></w:pict></w:r>`)}</w:hdr>`)).toEqual([]);
  });
  it("finds more than one, in order", () => {
    const xml = `${HDR_OPEN}${para(OLD_TEXT_WM + OLD_IMAGE_WM)}</w:hdr>`;
    expect(findWatermarks(xml).map((w) => w.kind)).toEqual(["text", "image"]);
  });
  it("markers are the ones Word uses", () => {
    expect(TEXT_MARKER).toBe("PowerPlusWaterMarkObject");
    expect(IMAGE_MARKER).toBe("WordPictureWatermark");
  });
});

describe("removing watermarks", () => {
  it("removes the whole run and keeps the other runs, text and paragraph properties", () => {
    const xml = `${HDR_OPEN}${para(textRun("Company") + OLD_TEXT_WM + textRun("Page"))}</w:hdr>`;
    const r = removeWatermarks(xml);
    expect(r.removed).toBe(1);
    expect(r.xml).toContain(`<w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:t>Company</w:t></w:r><w:r><w:t>Page</w:t></w:r></w:p>`);
    expect(findWatermarks(r.xml)).toEqual([]);
  });
  it("removes text and image watermarks together", () => {
    const r = removeWatermarks(`${HDR_OPEN}${para(OLD_TEXT_WM + OLD_IMAGE_WM)}</w:hdr>`);
    expect(r.removed).toBe(2);
  });
  it("a header without a watermark comes out unchanged", () => {
    const xml = `${HDR_OPEN}${para(textRun("Hello &amp; bye"))}</w:hdr>`;
    const r = removeWatermarks(xml);
    expect(r.removed).toBe(0);
    expect(r.xml).toBe(xml);
  });
  it("keeps a run with a non-watermark picture", () => {
    const xml = `${HDR_OPEN}${para(`<w:r><w:pict><v:shape id="Logo1"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:r>`)}</w:hdr>`;
    expect(removeWatermarks(xml).removed).toBe(0);
  });
});

describe("building a text watermark", () => {
  it("formatOpacity matches Word", () => {
    expect(formatOpacity(0.5)).toBe(".5");
    expect(formatOpacity(1)).toBe("1");
    expect(formatOpacity(0.25)).toBe(".25");
    expect(formatOpacity(0.333)).toBe(".33");
  });
  it("adds a diagonal text watermark to the first paragraph, after its properties", () => {
    const xml = `${HDR_OPEN}${para(textRun("Page"))}</w:hdr>`;
    const r = setWatermark(xml, "word/header1.xml", opts({ text: "DRAFT", color: "ff0000", opacity: 0.4 }));
    expect(r.added).toBe(true);
    expect(r.xml).toContain(`<w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>`);
    expect(r.xml).toContain(`string="DRAFT"`);
    expect(r.xml).toContain(`fillcolor="#FF0000"`);
    expect(r.xml).toContain(`<v:fill opacity=".4"/>`);
    expect(r.xml).toContain("rotation:315");
    expect(r.xml).toContain(`id="${TEXT_MARKER}100000001"`);
    expect(r.xml.indexOf("PowerPlus")).toBeLessThan(r.xml.indexOf("Page"));
    const found = findWatermarks(r.xml);
    expect(found).toHaveLength(1);
    expect(found[0]).toMatchObject({ kind: "text", text: "DRAFT", color: "#FF0000", opacity: ".4", rotated: true });
  });
  it("horizontal text has no rotation and a flatter box", () => {
    const r = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({ diagonal: false, text: "SAMPLE" }));
    expect(r.xml).not.toContain("rotation");
    expect(findWatermarks(r.xml)[0]?.rotated).toBe(false);
    expect(r.xml).toMatch(/width:420pt;height:105pt/);
  });
  it("size follows the text length within limits", () => {
    const small = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({ text: "A" })).xml;
    const big = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({ text: "A".repeat(40) })).xml;
    expect(small).toMatch(/width:150pt;height:90pt/);
    expect(big).toMatch(/width:520pt;height:312pt/);
  });
  it("escapes text and font name", () => {
    const r = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({ text: `Q&A <1> "x"`, fontFamily: `Arial "Narrow"` }));
    expect(r.xml).toContain(`string="Q&amp;A &lt;1&gt; &quot;x&quot;"`);
    expect(findWatermarks(r.xml)[0]?.text).toBe(`Q&A <1> "x"`);
    expect(() => rewriteXmlString(r.xml, "again.xml")).not.toThrow();
  });
  it("gives each header part its own shape ids through seed", () => {
    const a = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({ seed: 1 })).xml;
    const b = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({ seed: 2 })).xml;
    expect(findWatermarks(a)[0]?.id).not.toBe(findWatermarks(b)[0]?.id);
    expect(a).toContain("_x0000_s2049");
    expect(b).toContain("_x0000_s2050");
  });
  it("rejects bad values with INVALID_OPTIONS", () => {
    const bad: Partial<WatermarkOptions>[] = [
      { text: "" },
      { text: "   " },
      { text: "x".repeat(101) },
      { color: "red" },
      { color: "12345" },
      { opacity: 0 },
      { opacity: 1.5 },
      { opacity: Number.NaN },
      { fontFamily: " " },
      { seed: 0 },
      { seed: 1.5 },
    ];
    for (const over of bad) expect(codeOf(() => buildTextWatermarkRun(opts(over)))).toBe("INVALID_OPTIONS");
    expect(codeOf(() => buildTextWatermarkRun(opts({})))).toBe("NONE");
  });
});

describe("where the run goes", () => {
  it("into a paragraph without properties, before its first run", () => {
    const r = setWatermark(`${HDR_OPEN}<w:p>${textRun("Page")}</w:p></w:hdr>`, "h.xml", opts({}));
    expect(r.xml).toMatch(/<w:p><w:r><w:rPr><w:noProof\/><\/w:rPr><w:pict>/);
  });
  it("into an empty paragraph and an empty properties element", () => {
    expect(setWatermark(`${HDR_OPEN}<w:p/></w:hdr>`, "h.xml", opts({})).added).toBe(true);
    const withEmptyPPr = setWatermark(`${HDR_OPEN}<w:p><w:pPr/></w:p></w:hdr>`, "h.xml", opts({})).xml;
    expect(withEmptyPPr).toMatch(/<w:p><w:pPr\/><w:r>/);
  });
  it("only the first paragraph gets it", () => {
    const r = setWatermark(`${HDR_OPEN}${para(textRun("A"))}${para(textRun("B"))}</w:hdr>`, "h.xml", opts({}));
    expect(findWatermarks(r.xml)).toHaveLength(1);
    expect(r.xml.indexOf("PowerPlus")).toBeLessThan(r.xml.indexOf(">B<"));
  });
  it("a header that starts with a table puts it into the first top-level paragraph", () => {
    const xml = `${HDR_OPEN}<w:tbl><w:tr><w:tc><w:p>${textRun("cell")}</w:p></w:tc></w:tr></w:tbl>${para(textRun("after"))}</w:hdr>`;
    const r = setWatermark(xml, "h.xml", opts({}));
    expect(r.xml.indexOf("PowerPlus")).toBeGreaterThan(r.xml.indexOf("</w:tbl>"));
    expect(findWatermarks(r.xml)).toHaveLength(1);
  });
  it("a header without any paragraph gets one", () => {
    const r = setWatermark(`${HDR_OPEN}</w:hdr>`, "h.xml", opts({}));
    expect(r.added).toBe(true);
    expect(r.xml).toMatch(/<w:p><w:r><w:rPr><w:noProof\/><\/w:rPr><w:pict>[\s\S]*<\/w:pict><\/w:r><\/w:p><\/w:hdr>$/);
    expect(setWatermark(`<w:hdr xmlns:w="${W}"/>`, "h.xml", opts({})).added).toBe(true);
  });
  it("adds missing namespace declarations to the root, and no duplicates", () => {
    const r = setWatermark(`${BARE_OPEN}${para("")}</w:hdr>`, "h.xml", opts({}));
    for (const ns of ["xmlns:v=", "xmlns:o=", "xmlns:w10=", "xmlns:r="]) expect(r.xml.split(ns).length - 1).toBe(1);
    const again = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", opts({}));
    expect(again.xml.split("xmlns:v=").length - 1).toBe(1);
  });
  it("createRunInserter reports whether it inserted", () => {
    const inserter = createRunInserter(buildTextWatermarkRun(opts({})));
    const rewriter = createXmlRewriter({ partName: "h.xml", onOutput: () => undefined, transform: inserter.transform });
    rewriter.write(`${HDR_OPEN}${para("")}</w:hdr>`);
    rewriter.end();
    expect(inserter.stats.inserted).toBe(true);
  });
});

describe("modes", () => {
  const withOld = `${HDR_OPEN}${para(textRun("Page") + OLD_TEXT_WM)}</w:hdr>`;
  it("inherit and keepSource leave the header exactly as it is", () => {
    for (const mode of ["inherit", "keepSource"] as const) {
      expect(setWatermark(withOld, "h.xml", { mode })).toEqual({ xml: withOld, removed: 0, added: false });
    }
  });
  it("remove removes", () => {
    const r = setWatermark(withOld, "h.xml", { mode: "remove" });
    expect(r).toMatchObject({ removed: 1, added: false });
    expect(findWatermarks(r.xml)).toEqual([]);
    expect(r.xml).toContain("Page");
  });
  it("text replaces an old watermark: exactly one afterwards", () => {
    const r = setWatermark(withOld, "h.xml", opts({ text: "NEW" }));
    expect(r).toMatchObject({ removed: 1, added: true });
    const found = findWatermarks(r.xml);
    expect(found).toHaveLength(1);
    expect(found[0]?.text).toBe("NEW");
  });
  it("running text twice is stable and still one watermark", () => {
    const once = setWatermark(withOld, "h.xml", opts({ text: "NEW" })).xml;
    const twice = setWatermark(once, "h.xml", opts({ text: "NEW" })).xml;
    expect(twice).toBe(once);
  });
  it("the result is well formed XML", () => {
    for (const mode of ["text", "remove"] as const) {
      const r = setWatermark(withOld, "h.xml", opts({ mode }));
      expect(() => rewriteXmlString(r.xml, "again.xml")).not.toThrow();
    }
  });
  it("damaged XML throws ZIP_READ_FAILED", () => {
    expect(codeOf(() => setWatermark("<w:hdr><w:p>", "h.xml", opts({})))).toBe("ZIP_READ_FAILED");
    expect(codeOf(() => setWatermark("<w:hdr><w:p>", "h.xml", { mode: "remove" }))).toBe("ZIP_READ_FAILED");
  });
});

describe("image watermark", () => {
  const image = (over: Partial<WatermarkOptions> = {}): WatermarkOptions => ({ mode: "image", imageRelId: "rIdWm1", widthPt: 300, heightPt: 200, ...over });
  it("adds a picture shape that points at the relationship id", () => {
    const r = setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", image());
    expect(r.added).toBe(true);
    const found = findWatermarks(r.xml);
    expect(found).toHaveLength(1);
    expect(found[0]).toMatchObject({ kind: "image", relId: "rIdWm1", id: `${IMAGE_MARKER}100000001` });
    expect(r.xml).toContain(`gain="19661f"`);
    expect(r.xml).toContain("width:300pt;height:200pt");
  });
  it("washout can be turned off", () => {
    expect(setWatermark(`${HDR_OPEN}${para("")}</w:hdr>`, "h.xml", image({ washout: false })).xml).not.toContain("gain=");
  });
  it("replaces an old text watermark with the picture", () => {
    const r = setWatermark(`${HDR_OPEN}${para(OLD_TEXT_WM)}</w:hdr>`, "h.xml", image());
    expect(r.removed).toBe(1);
    expect(findWatermarks(r.xml).map((w) => w.kind)).toEqual(["image"]);
  });
  it("rejects missing or bad picture settings", () => {
    const bad: Partial<WatermarkOptions>[] = [{ imageRelId: undefined }, { imageRelId: " " }, { widthPt: undefined }, { heightPt: 0 }, { widthPt: 1501 }, { heightPt: -5 }, { seed: 0 }];
    for (const over of bad) expect(codeOf(() => buildImageWatermarkRun(image(over)))).toBe("INVALID_OPTIONS");
  });
});

describe("addImageRelationship", () => {
  it("creates a relationships file when there is none", () => {
    const r = addImageRelationship(undefined, "media/wm.png");
    expect(r.id).toBe("rIdWm1");
    expect(parseRelationships(r.xml)).toEqual([
      { id: "rIdWm1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: "media/wm.png" },
    ]);
  });
  it("keeps existing relationships and picks the first free id", () => {
    const existing = addImageRelationship(undefined, "media/a.png").xml;
    const r = addImageRelationship(existing, "media/b.png");
    expect(r.id).toBe("rIdWm2");
    expect(parseRelationships(r.xml).map((x) => x.id)).toEqual(["rIdWm1", "rIdWm2"]);
  });
});
