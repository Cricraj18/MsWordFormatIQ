import { beforeAll, describe, expect, it } from "vitest";
import { Document, Footer, Header, HeadingLevel, Packer, PageOrientation, Paragraph, TableOfContents, TextRun } from "docx";
import { DEFAULT_PAGE, extractTemplateInfo } from "./template_extract";
import { createBlobZipOutput, createJobControl, setWorkersEnabled } from "./zipio";

beforeAll(() => setWorkersEnabled(false));

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const STYLES = `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>
<w:latentStyles w:count="1"><w:lsdException w:name="Normal"/></w:latentStyles>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
<w:style w:type="numbering" w:styleId="NoList"><w:name w:val="No List"/></w:style>
<w:style w:type="other" w:styleId="Odd"><w:name w:val="Odd"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="duplicate"/></w:style>
</w:styles>`;

function doc(body: string): string {
  return `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`;
}

async function build(opts: { styles?: string | null; document?: string; extra?: Record<string, string> }): Promise<Blob> {
  const { output, getBlob } = createBlobZipOutput();
  await output.addText("[Content_Types].xml", "<Types/>");
  await output.addText("word/document.xml", opts.document ?? doc("<w:p/>"));
  if (opts.styles !== null) await output.addText("word/styles.xml", opts.styles ?? STYLES);
  for (const [path, text] of Object.entries(opts.extra ?? {})) await output.addText(path, text);
  await output.finish();
  return getBlob();
}

const opts = { id: "t1", name: "Test template", createdAt: "2026-10-11T00:00:00Z" };

describe("styles", () => {
  it("reads kinds and names, skips unknown types, keeps the first duplicate", async () => {
    const info = await extractTemplateInfo(await build({}), "t.docx", opts);
    expect(info.styles).toEqual([
      { id: "Normal", name: "Normal", kind: "paragraph" },
      { id: "Heading1", name: "heading 1", kind: "paragraph" },
      { id: "Strong", name: "Strong", kind: "character" },
      { id: "TableGrid", name: "Table Grid", kind: "table" },
      { id: "NoList", name: "No List", kind: "numbering" },
    ]);
    expect(info.id).toBe("t1");
    expect(info.name).toBe("Test template");
    expect(info.createdAt).toBe("2026-10-11T00:00:00Z");
  });

  it("works with a different namespace prefix", async () => {
    const styles = `<x:styles xmlns:x="u"><x:style x:type="paragraph" x:styleId="A"><x:name x:val="Alpha"/></x:style></x:styles>`;
    const info = await extractTemplateInfo(await build({ styles }), "t.docx", opts);
    expect(info.styles).toEqual([{ id: "A", name: "Alpha", kind: "paragraph" }]);
  });

  it("fails with TEMPLATE_INVALID when styles.xml is missing", async () => {
    await expect(extractTemplateInfo(await build({ styles: null }), "t.docx", opts)).rejects.toMatchObject({ code: "TEMPLATE_INVALID" });
  });

  it("handles a template with no styles at all", async () => {
    const info = await extractTemplateInfo(await build({ styles: `<w:styles ${W}/>` }), "t.docx", opts);
    expect(info.styles).toEqual([]);
  });
});

describe("page setup", () => {
  it("uses A4 portrait when there is no sectPr", async () => {
    const info = await extractTemplateInfo(await build({}), "t.docx", opts);
    expect(info.page).toEqual(DEFAULT_PAGE);
  });

  it("reads portrait Letter", async () => {
    const d = doc('<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>');
    const info = await extractTemplateInfo(await build({ document: d }), "t.docx", opts);
    expect(info.page).toEqual({ widthTwips: 12240, heightTwips: 15840, landscape: false });
  });

  it("reads landscape from the orient attribute", async () => {
    const d = doc('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>');
    const info = await extractTemplateInfo(await build({ document: d }), "t.docx", opts);
    expect(info.page).toEqual({ widthTwips: 16838, heightTwips: 11906, landscape: true });
  });

  it("uses the FIRST sectPr, not the last", async () => {
    const d = doc('<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>');
    const info = await extractTemplateInfo(await build({ document: d }), "t.docx", opts);
    expect(info.page.widthTwips).toBe(12240);
  });

  it("ignores a pgSz inside sectPrChange (old tracked value)", async () => {
    const d = doc('<w:sectPr><w:sectPrChange w:id="1"><w:sectPr><w:pgSz w:w="1000" w:h="2000"/></w:sectPr></w:sectPrChange><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>');
    const info = await extractTemplateInfo(await build({ document: d }), "t.docx", opts);
    expect(info.page).toEqual({ widthTwips: 11906, heightTwips: 16838, landscape: false });
  });

  it("falls back to A4 when pgSz numbers are bad", async () => {
    const d = doc('<w:sectPr><w:pgSz w:w="abc" w:h="-5"/></w:sectPr>');
    const info = await extractTemplateInfo(await build({ document: d }), "t.docx", opts);
    expect(info.page).toEqual(DEFAULT_PAGE);
  });
});

describe("header, footer and TOC", () => {
  it("hasHeaderFooter is false without header/footer parts", async () => {
    expect((await extractTemplateInfo(await build({}), "t.docx", opts)).hasHeaderFooter).toBe(false);
  });

  it("hasHeaderFooter is true with a header part or a footer part", async () => {
    const a = await extractTemplateInfo(await build({ extra: { "word/header1.xml": "<w:hdr/>" } }), "t.docx", opts);
    const b = await extractTemplateInfo(await build({ extra: { "word/footer2.xml": "<w:ftr/>" } }), "t.docx", opts);
    expect(a.hasHeaderFooter).toBe(true);
    expect(b.hasHeaderFooter).toBe(true);
  });

  it("finds a TOC field written as instrText", async () => {
    const d = doc('<w:p><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r></w:p>');
    expect((await extractTemplateInfo(await build({ document: d }), "t.docx", opts)).hasToc).toBe(true);
  });

  it("finds a TOC in a second instrText after a blank one", async () => {
    const d = doc("<w:p><w:r><w:instrText> </w:instrText><w:instrText>TOC</w:instrText></w:r></w:p>");
    expect((await extractTemplateInfo(await build({ document: d }), "t.docx", opts)).hasToc).toBe(true);
  });

  it("finds a TOC written as fldSimple", async () => {
    const d = doc('<w:fldSimple w:instr=" TOC \\o &quot;1-3&quot; "><w:r/></w:fldSimple>');
    expect((await extractTemplateInfo(await build({ document: d }), "t.docx", opts)).hasToc).toBe(true);
  });

  it("finds a Table of Contents content control", async () => {
    const d = doc('<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr></w:sdt>');
    expect((await extractTemplateInfo(await build({ document: d }), "t.docx", opts)).hasToc).toBe(true);
  });

  it("does not treat other fields as a TOC", async () => {
    const d = doc("<w:p><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:instrText> TOCX </w:instrText></w:r></w:p>");
    expect((await extractTemplateInfo(await build({ document: d }), "t.docx", opts)).hasToc).toBe(false);
  });
});

describe("errors, safety and control", () => {
  it("rejects a wrong extension with NOT_DOCX", async () => {
    await expect(extractTemplateInfo(await build({}), "t.pdf", opts)).rejects.toMatchObject({ code: "NOT_DOCX" });
  });

  it("rejects text that is not a zip", async () => {
    await expect(extractTemplateInfo(new Blob(["hello"]), "t.docx", opts)).rejects.toMatchObject({ code: "ZIP_READ_FAILED" });
  });

  it("rejects a DOCTYPE entity bomb in styles.xml", async () => {
    const styles = '<!DOCTYPE s [<!ENTITY a "lol"><!ENTITY b "&a;&a;">]><w:styles xmlns:w="x">&b;</w:styles>';
    await expect(extractTemplateInfo(await build({ styles }), "t.docx", opts)).rejects.toMatchObject({ code: "XML_ENTITY_EXPANSION" });
  });

  it("rejects a DOCTYPE entity bomb in document.xml", async () => {
    const document = '<!DOCTYPE s [<!ENTITY a "lol">]><w:document xmlns:w="x">&a;</w:document>';
    await expect(extractTemplateInfo(await build({ document }), "t.docx", opts)).rejects.toMatchObject({ code: "XML_ENTITY_EXPANSION" });
  });

  it("rejects damaged XML in document.xml", async () => {
    await expect(extractTemplateInfo(await build({ document: "<w:document><w:body>" }), "t.docx", opts)).rejects.toMatchObject({ code: "ZIP_READ_FAILED" });
  });

  it("stops with CANCELLED when the job is already cancelled", async () => {
    const ac = new AbortController();
    ac.abort();
    const control = createJobControl(60_000, ac.signal);
    await expect(extractTemplateInfo(await build({}), "t.docx", { ...opts, control })).rejects.toMatchObject({ code: "CANCELLED" });
    control.dispose();
  });

  it("reports progress from 0 to 1", async () => {
    const fractions: number[] = [];
    await extractTemplateInfo(await build({}), "t.docx", { ...opts, onProgress: (p) => fractions.push(p.fraction) });
    expect(fractions[0]).toBe(0);
    expect(fractions[fractions.length - 1]).toBe(1);
    expect([...fractions].sort((a, b) => a - b)).toEqual(fractions);
  });

  it("does not change the source blob", async () => {
    const blob = await build({});
    const before = new Uint8Array(await blob.arrayBuffer());
    await extractTemplateInfo(blob, "t.docx", opts);
    const after = new Uint8Array(await blob.arrayBuffer());
    expect(after).toEqual(before);
  });

  it("reads a 60000-paragraph document.xml (about 8 MB) quickly", async () => {
    const body = '<w:p><w:r><w:t>Some paragraph text here</w:t></w:r></w:p>'.repeat(60_000);
    const d = doc(`${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`);
    const t0 = Date.now();
    const info = await extractTemplateInfo(await build({ document: d }), "t.docx", opts);
    expect(info.page.widthTwips).toBe(11906);
    expect(Date.now() - t0).toBeLessThan(10_000);
  });
});

describe("real-style .docx made by the docx library", () => {
  it("reads headings, TOC, header, footer and landscape from it", async () => {
    const file = new Document({
      features: { updateFields: true },
      sections: [
        {
          properties: { page: { size: { orientation: PageOrientation.LANDSCAPE } } },
          headers: { default: new Header({ children: [new Paragraph("Header text")] }) },
          footers: { default: new Footer({ children: [new Paragraph("Footer text")] }) },
          children: [
            new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-3" }),
            new Paragraph({ text: "1 Introduction", heading: HeadingLevel.HEADING_1 }),
            new Paragraph({ children: [new TextRun("Body text.")] }),
          ],
        },
      ],
    });
    const buffer = await Packer.toBuffer(file);
    const info = await extractTemplateInfo(new Blob([new Uint8Array(buffer)]), "real.docx", opts);
    expect(info.hasHeaderFooter).toBe(true);
    expect(info.hasToc).toBe(true);
    expect(info.page.landscape).toBe(true);
    expect(info.page.widthTwips).toBeGreaterThan(info.page.heightTwips);
    expect(info.styles.some((s) => s.id === "Heading1" && s.kind === "paragraph")).toBe(true);
  });
});
