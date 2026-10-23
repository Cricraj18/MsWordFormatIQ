import { describe, expect, it } from "vitest";
import { createXmlFeeder } from "./xml";
import { createParagraphHandlers, createStyleHandlers, resolveStyleHints, MAX_PARAGRAPH_TEXT } from "./paragraphs";
import type { ParagraphFacts } from "./paragraphs";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function readPs(body: string): ParagraphFacts[] {
  const out: ParagraphFacts[] = [];
  const feeder = createXmlFeeder(createParagraphHandlers((f) => out.push(f)), "document.xml");
  feeder.write(`<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  feeder.end();
  return out;
}
const run = (t: string, rpr = ""): string => `<w:r>${rpr === "" ? "" : `<w:rPr>${rpr}</w:rPr>`}<w:t xml:space="preserve">${t}</w:t></w:r>`;

describe("paragraph reader", () => {
  it("reads text, index and style", () => {
    const [a, b] = readPs(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run("Hello ")}${run("world")}</w:p><w:p>${run("Second")}</w:p>`);
    expect(a).toMatchObject({ index: 0, text: "Hello world", styleId: "Heading1", inTable: false, hasImage: false, isListItem: false });
    expect(b).toMatchObject({ index: 1, text: "Second", styleId: null });
  });
  it("handles an empty paragraph and self-closing p", () => {
    const list = readPs("<w:p/><w:p></w:p>");
    expect(list.map((p) => p.text)).toEqual(["", ""]);
  });
  it("marks paragraphs in tables, also nested tables", () => {
    const list = readPs(`<w:p>${run("a")}</w:p><w:tbl><w:tr><w:tc><w:p>${run("b")}</w:p><w:tbl><w:tr><w:tc><w:p>${run("c")}</w:p></w:tc></w:tr></w:tbl><w:p>${run("d")}</w:p></w:tc></w:tr></w:tbl><w:p>${run("e")}</w:p>`);
    expect(list.map((p) => [p.text, p.inTable])).toEqual([["a", false], ["b", true], ["c", true], ["d", true], ["e", false]]);
  });
  it("detects list numbering but not numId 0", () => {
    const list = readPs(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>${run("x")}</w:p><w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>${run("y")}</w:p>`);
    expect(list.map((p) => p.isListItem)).toEqual([true, false]);
  });
  it("reads paragraph outline level (0 based in XML, 1 based here); 9 means none", () => {
    const list = readPs(`<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr>${run("x")}</w:p><w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr>${run("y")}</w:p>`);
    expect(list.map((p) => p.outlineLevel)).toEqual([2, null]);
  });
  it("sees drawings and pictures", () => {
    const list = readPs(`<w:p><w:r><w:drawing/></w:r></w:p><w:p><w:r><w:pict/></w:r></w:p><w:p>${run("t")}</w:p>`);
    expect(list.map((p) => p.hasImage)).toEqual([true, true, false]);
  });
  it("ignores deleted text and field codes, keeps inserted text, tab and break", () => {
    const [p] = readPs(`<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r><w:del><w:r><w:delText>GONE</w:delText></w:r></w:del><w:ins><w:r><w:t>NEW</w:t></w:r></w:ins><w:r><w:instrText> TOC \\o </w:instrText></w:r></w:p>`);
    expect(p?.text).toBe("A\tB CNEW");
  });
  it("ignores tab stops inside pPr", () => {
    const [p] = readPs(`<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>${run("x")}</w:p>`);
    expect(p?.text).toBe("x");
  });
  it("skips nested paragraphs (text boxes) and keeps outer text", () => {
    const list = readPs(`<w:p>${run("outer")}<w:r><w:txbxContent><w:p>${run("inner")}</w:p></w:txbxContent></w:r>${run(" more")}</w:p><w:p>${run("next")}</w:p>`);
    expect(list.map((p) => p.text)).toEqual(["outer more", "next"]);
    expect(list.map((p) => p.index)).toEqual([0, 1]);
  });
  it("cuts the text at MAX_PARAGRAPH_TEXT characters", () => {
    const [p] = readPs(`<w:p>${run("x".repeat(1000))}</w:p>`);
    expect(p?.text.length).toBe(MAX_PARAGRAPH_TEXT);
  });
  it("bold: all bold, none, mixed, explicit off, whitespace-only runs ignored", () => {
    const list = readPs(
      `<w:p>${run("a", "<w:b/>")}${run("b", '<w:b w:val="1"/>')}</w:p>` +
      `<w:p>${run("a")}</w:p>` +
      `<w:p>${run("a", "<w:b/>")}${run("b")}</w:p>` +
      `<w:p>${run("a", '<w:b w:val="0"/>')}</w:p>` +
      `<w:p>${run("a", "<w:b/>")}${run("  ")}</w:p>` +
      `<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr>${run("a")}</w:p>`,
    );
    expect(list.map((p) => p.bold)).toEqual([true, null, false, false, true, null]);
  });
  it("font size: biggest run with text wins; run without text ignored", () => {
    const list = readPs(
      `<w:p>${run("a", '<w:sz w:val="28"/>')}${run("b", '<w:sz w:val="40"/>')}</w:p>` +
      `<w:p>${run("a")}</w:p>` +
      `<w:p>${run("a", '<w:sz w:val="24"/>')}<w:r><w:rPr><w:sz w:val="80"/></w:rPr></w:r></w:p>`,
    );
    expect(list.map((p) => p.sizeHalfPoints)).toEqual([40, null, 24]);
  });
  it("works with another namespace prefix", () => {
    const out: ParagraphFacts[] = [];
    const f = createXmlFeeder(createParagraphHandlers((x) => out.push(x)), "d");
    f.write('<x:document xmlns:x="u"><x:body><x:p><x:r><x:t>Hi</x:t></x:r></x:p></x:body></x:document>');
    f.end();
    expect(out[0]?.text).toBe("Hi");
  });
  it("is fed one character at a time with the same result", () => {
    const xml = `<w:document ${NS}><w:body><w:p><w:pPr><w:pStyle w:val="H"/></w:pPr>${run("Grüße 日本 😀", "<w:b/>")}</w:p></w:body></w:document>`;
    const out: ParagraphFacts[] = [];
    const f = createXmlFeeder(createParagraphHandlers((x) => out.push(x)), "d");
    for (const ch of xml) f.write(ch);
    f.end();
    expect(out[0]).toMatchObject({ text: "Grüße 日本 😀", styleId: "H", bold: true });
  });
  it("handles 100000 paragraphs without keeping them", () => {
    let count = 0;
    let last = -1;
    const f = createXmlFeeder(createParagraphHandlers((x) => { count += 1; last = x.index; }), "d");
    f.write(`<w:document ${NS}><w:body>`);
    const piece = `<w:p>${run("some text here")}</w:p>`.repeat(1000);
    for (let i = 0; i < 100; i += 1) f.write(piece);
    f.write("</w:body></w:document>");
    f.end();
    expect(count).toBe(100000);
    expect(last).toBe(99999);
  });
});

function readStyles(xml: string): ReturnType<typeof resolveStyleHints> {
  const raw: Parameters<typeof createStyleHandlers>[0] = new Map();
  const defaults: { size: number | null } = { size: null };
  const f = createXmlFeeder(createStyleHandlers(raw, defaults), "styles.xml");
  f.write(`<w:styles ${NS}>${xml}</w:styles>`);
  f.end();
  return resolveStyleHints(raw, defaults.size);
}

describe("style hints", () => {
  it("reads name, outline level, bold, size, and default size", () => {
    const h = readStyles(
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>`,
    );
    expect(h.defaultSizeHalfPoints).toBe(22);
    expect(h.styles.get("Heading1")).toEqual({ name: "heading 1", outlineLevel: 1, bold: true, sizeHalfPoints: 32 });
    expect(h.styles.has("Strong")).toBe(false);
  });
  it("follows basedOn, own values win, loops are safe", () => {
    const h = readStyles(
      `<w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:rPr><w:b/><w:sz w:val="30"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="L1"><w:name w:val="L1"/><w:basedOn w:val="L2"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="L2"><w:name w:val="L2"/><w:basedOn w:val="L1"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>`,
    );
    expect(h.styles.get("B")).toMatchObject({ bold: true, sizeHalfPoints: 20 });
    expect(h.styles.get("L1")).toMatchObject({ bold: false });
  });
  it("first of duplicate ids wins; missing name falls back to the id", () => {
    const h = readStyles(
      `<w:style w:type="paragraph" w:styleId="X"><w:name w:val="First"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="X"><w:name w:val="Second"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Y"></w:style>`,
    );
    expect(h.styles.get("X")?.name).toBe("First");
    expect(h.styles.get("Y")?.name).toBe("Y");
  });
});
