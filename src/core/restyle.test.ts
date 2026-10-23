import { describe, expect, it } from "vitest";
import { createXmlFeeder } from "./xml";
import { createParagraphHandlers } from "./paragraphs";
import type { ParagraphFacts } from "./paragraphs";
import { createRestyler } from "./restyle";
import type { DirectFormatting } from "./restyle";
import { createXmlRewriter, rewriteXmlString } from "./xml_rewrite";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const doc = (body: string): string => `<w:document ${NS}><w:body>${body}</w:body></w:document>`;

function run(xml: string, styleFor: (i: number) => string | null, mode: DirectFormatting = "clear") {
  const restyler = createRestyler({ styleFor, directFormatting: mode });
  const out = rewriteXmlString(xml, "document.xml", restyler.transform);
  return { out, stats: restyler.stats };
}

function readFacts(xml: string): ParagraphFacts[] {
  const list: ParagraphFacts[] = [];
  const feeder = createXmlFeeder(createParagraphHandlers((f) => list.push(f)), "t.xml");
  feeder.write(xml);
  feeder.end();
  return list;
}

describe("setting the style", () => {
  it("creates pPr and pStyle for a paragraph without pPr", () => {
    const { out } = run(doc(`<w:p><w:r><w:t>Hi</w:t></w:r></w:p>`), () => "Heading1");
    expect(out).toContain(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`);
  });
  it("replaces an existing pStyle and puts pStyle first", () => {
    const { out } = run(doc(`<w:p><w:pPr><w:keepNext/><w:pStyle w:val="Old"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`), () => "New");
    expect(out).toContain(`<w:pPr><w:pStyle w:val="New"/><w:keepNext/></w:pPr>`);
    expect(out).not.toContain("Old");
  });
  it("adds pStyle first inside a pPr that has none", () => {
    const { out } = run(doc(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`), () => "ListParagraph");
    expect(out).toContain(`<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr>`);
    expect(out).toContain(`<w:numId w:val="3"/>`);
  });
  it("gives an empty self-closing paragraph a pPr", () => {
    const { out } = run(doc(`<w:p/>`), () => "Normal");
    expect(out).toContain(`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>`);
  });
  it("does not duplicate pPr when whitespace sits before it", () => {
    const { out } = run(doc(`<w:p>\n  <w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`), () => "B");
    expect(out.match(/<w:pPr>/g)?.length).toBe(1);
    expect(out).toContain(`<w:pStyle w:val="B"/>`);
  });
  it("keeps the section break inside pPr and puts pStyle before it", () => {
    const sect = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440"/></w:sectPr>`;
    const { out } = run(doc(`<w:p><w:pPr>${sect}</w:pPr></w:p>`), () => "Normal");
    expect(out).toContain(`<w:pPr><w:pStyle w:val="Normal"/>${sect}</w:pPr>`);
  });
  it("escapes odd characters in the style id", () => {
    const { out } = run(doc(`<w:p/>`), () => 'A&B"C');
    expect(out).toContain(`w:val="A&amp;B&quot;C"`);
  });
  it("uses the same prefix as the document", () => {
    const xml = `<x:document xmlns:x="urn:w"><x:body><x:p><x:r><x:t>a</x:t></x:r></x:p></x:body></x:document>`;
    const { out } = run(xml, () => "S");
    expect(out).toContain(`<x:p><x:pPr><x:pStyle x:val="S"/></x:pPr><x:r>`);
  });
  it("is stable when run twice with the same styles", () => {
    const src = doc(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="40"/><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>`);
    const once = run(src, () => "H").out;
    expect(run(once, () => "H").out).toBe(once);
  });
});

describe("leaving paragraphs alone", () => {
  it("copies a paragraph exactly when styleFor gives null", () => {
    const body = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="40"/></w:rPr><w:t>x</w:t></w:r></w:p>`;
    const { out, stats } = run(doc(body), () => null);
    expect(out).toBe(doc(body));
    expect(stats.restyled).toBe(0);
    expect(stats.paragraphs).toBe(1);
  });
  it("only some paragraphs are changed", () => {
    const { out } = run(doc(`<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p>`), (i) => (i === 1 ? "X" : null));
    expect(out).toContain(`<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="X"/></w:pPr>`);
  });
});

describe("clearing direct formatting", () => {
  const messy =
    `<w:p><w:pPr><w:keepNext/><w:numPr><w:numId w:val="2"/></w:numPr><w:pBdr><w:bottom w:val="single"/></w:pBdr><w:shd w:fill="EEEEEE"/>` +
    `<w:spacing w:before="240"/><w:ind w:left="720"/><w:contextualSpacing/><w:jc w:val="center"/><w:outlineLvl w:val="0"/>` +
    `<w:rPr><w:rFonts w:ascii="Arial"/><w:b/><w:sz w:val="48"/></w:rPr></w:pPr>` +
    `<w:r><w:rPr><w:rStyle w:val="Strong"/><w:rFonts w:ascii="Comic"/><w:b/><w:i/><w:u w:val="single"/><w:sz w:val="40"/><w:szCs w:val="40"/>` +
    `<w:color w:val="FF0000"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/><w:spacing w:val="20"/><w:kern w:val="2"/><w:lang w:val="en-GB"/></w:rPr><w:t>Text</w:t></w:r></w:p>`;
  it("removes paragraph and run formatting that hides the style", () => {
    const { out, stats } = run(doc(messy), () => "Heading1", "clear");
    for (const gone of ["w:jc", "w:spacing", "w:ind", "w:contextualSpacing", "w:shd", "w:pBdr", "w:rFonts", "w:sz ", "w:szCs", "w:color", "w:kern"]) {
      expect(out).not.toContain("<" + gone);
    }
    expect(stats.clearedParagraphProps).toBe(6);
    expect(stats.clearedRunProps).toBe(8);
  });
  it("keeps list, layout, emphasis, language and section-relevant parts", () => {
    const { out } = run(doc(messy), () => "Heading1", "clear");
    for (const kept of ["<w:keepNext/>", '<w:numId w:val="2"/>', '<w:outlineLvl w:val="0"/>', '<w:rStyle w:val="Strong"/>', "<w:b/>", "<w:i/>", '<w:u w:val="single"/>', '<w:highlight w:val="yellow"/>', '<w:vertAlign w:val="superscript"/>', '<w:lang w:val="en-GB"/>']) {
      expect(out).toContain(kept);
    }
  });
  it("keep mode changes only the style", () => {
    const { out, stats } = run(doc(messy), () => "Heading1", "keep");
    expect(out).toContain("<w:jc");
    expect(out).toContain("<w:rFonts");
    expect(out).toContain('<w:pStyle w:val="Heading1"/>');
    expect(stats.clearedParagraphProps + stats.clearedRunProps).toBe(0);
  });
  it("does not touch tracked formatting change records", () => {
    const src = doc(
      `<w:p><w:pPr><w:jc w:val="left"/><w:pPrChange w:id="1"><w:pPr><w:jc w:val="center"/><w:spacing w:after="10"/></w:pPr></w:pPrChange></w:pPr>` +
        `<w:r><w:rPr><w:sz w:val="30"/><w:rPrChange w:id="2"><w:rPr><w:rFonts w:ascii="Old"/><w:sz w:val="20"/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>`,
    );
    const { out } = run(src, () => "Normal", "clear");
    expect(out).toContain(`<w:pPrChange w:id="1"><w:pPr><w:jc w:val="center"/><w:spacing w:after="10"/></w:pPr></w:pPrChange>`);
    expect(out).toContain(`<w:rPrChange w:id="2"><w:rPr><w:rFonts w:ascii="Old"/><w:sz w:val="20"/></w:rPr></w:rPrChange>`);
    expect(out).not.toContain(`<w:jc w:val="left"/>`);
    expect(out).not.toContain(`<w:sz w:val="30"/>`);
    expect(out.match(/<w:pStyle /g)?.length).toBe(1);
  });
  it("does not clear formatting in paragraphs that are left alone", () => {
    const body = `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:sz w:val="50"/></w:rPr><w:t>x</w:t></w:r></w:p>`;
    expect(run(doc(body), () => null, "clear").out).toBe(doc(body));
  });
});

describe("numbering of paragraphs", () => {
  it("nested text-box paragraphs are not counted and not changed", () => {
    const body =
      `<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>` +
      `<w:p><w:r><w:t>second</w:t></w:r></w:p>`;
    const seen: number[] = [];
    const { out, stats } = run(doc(body), (i) => {
      seen.push(i);
      return "S" + i;
    });
    expect(seen).toEqual([0, 1]);
    expect(stats.paragraphs).toBe(2);
    expect(out).toContain(`<w:txbxContent><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent>`);
    expect(out).toContain(`<w:pStyle w:val="S1"/>`);
  });
  it("paragraphs in tables and content controls count", () => {
    const body =
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
      `<w:sdt><w:sdtContent><w:p><w:r><w:t>s</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p/>`;
    const seen: number[] = [];
    run(doc(body), (i) => {
      seen.push(i);
      return null;
    });
    expect(seen).toEqual([0, 1, 2, 3]);
  });
});

describe("index parity with the paragraph reader (paragraphs.ts)", () => {
  const messyDoc = doc(
    `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>` +
      `<w:p/>` +
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p/></w:tc><w:tc><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>` +
      `<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData><wps:txbx><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p><w:p><w:r><w:t>box2</w:t></w:r></w:p></w:txbxContent></wps:txbx></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
      `<w:sdt><w:sdtContent><w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:pPr></w:p>` +
      `<w:p><w:r><w:t>last</w:t></w:r></w:p>`,
  );
  it("every paragraph the reader counts gets the style made for its own number", () => {
    const before = readFacts(messyDoc);
    const { out, stats } = run(messyDoc, (i) => "S" + i);
    const after = readFacts(out);
    expect(after).toHaveLength(before.length);
    expect(stats.paragraphs).toBe(before.length);
    expect(after.map((f) => f.styleId)).toEqual(before.map((f) => "S" + f.index));
    expect(after.map((f) => f.text)).toEqual(before.map((f) => f.text));
    expect(after.map((f) => f.inTable)).toEqual(before.map((f) => f.inTable));
    expect(after.map((f) => f.isListItem)).toEqual(before.map((f) => f.isListItem));
  });
  it("output is well formed and keeps the section break", () => {
    const { out } = run(messyDoc, (i) => "S" + i);
    expect(out).toContain(`<w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr>`);
    expect(() => rewriteXmlString(out, "again.xml")).not.toThrow();
  });
});

describe("streaming size", () => {
  it("restyles 100000 paragraphs fed in pieces, in order", () => {
    const parts: string[] = [`<w:body ${NS}>`];
    for (let i = 0; i < 100000; i += 1) {
      parts.push(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="40"/><w:b/></w:rPr><w:t>Line ${i}</w:t></w:r></w:p>`);
    }
    parts.push("</w:body>");
    const xml = parts.join("");
    const restyler = createRestyler({ styleFor: (i) => (i % 2 === 0 ? "A" : "B"), directFormatting: "clear" });
    let length = 0;
    const rewriter = createXmlRewriter({ partName: "big.xml", onOutput: (c) => { length += c.length; }, transform: restyler.transform });
    for (let i = 0; i < xml.length; i += 65536) rewriter.write(xml.slice(i, i + 65536));
    rewriter.end();
    expect(restyler.stats.paragraphs).toBe(100000);
    expect(restyler.stats.restyled).toBe(100000);
    expect(restyler.stats.clearedParagraphProps).toBe(100000);
    expect(restyler.stats.clearedRunProps).toBe(100000);
    expect(length).toBeGreaterThan(0);
    expect(length).toBeLessThan(xml.length);
  });
});
