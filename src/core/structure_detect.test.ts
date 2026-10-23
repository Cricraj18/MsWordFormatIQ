import { beforeAll, describe, expect, it } from "vitest";
import { Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx";
import { createBlobZipOutput, createJobControl, setWorkersEnabled } from "./zipio";
import { detectStructure, detectStructureStream } from "./structure_detect";
import { LOW_CONFIDENCE_THRESHOLD } from "../types";
import type { DetectedParagraph, ProgressInfo } from "../types";

beforeAll(() => setWorkersEnabled(false));

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const run = (t: string, rpr = ""): string => `<w:r>${rpr === "" ? "" : `<w:rPr>${rpr}</w:rPr>`}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const p = (t: string, ppr = "", rpr = ""): string => `<w:p>${ppr === "" ? "" : `<w:pPr>${ppr}</w:pPr>`}${run(t, rpr)}</w:p>`;

async function makeDocx(body: string, styles?: string, doctype = ""): Promise<Blob> {
  const { output, getBlob } = createBlobZipOutput();
  await output.addText("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  await output.addText("word/document.xml", `<?xml version="1.0"?>${doctype}<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  if (styles !== undefined) await output.addText("word/styles.xml", `<?xml version="1.0"?><w:styles ${NS}>${styles}</w:styles>`);
  await output.finish();
  return getBlob();
}

const STYLES =
  `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:sz w:val="52"/></w:rPr></w:style>`;

describe("detectStructure on hand-made files", () => {
  it("classifies a small mixed document", async () => {
    const body =
      p("My Report", '<w:pStyle w:val="Title"/>') +
      p("Chapter 1") +
      p("1.1 Background") +
      p("This is a normal body paragraph with enough words to look like real text in a document.") +
      p("Figure 1: The plan") +
      p("Item one", '<w:numPr><w:numId w:val="2"/></w:numPr>') +
      `<w:tbl><w:tr><w:tc>${p("cell")}</w:tc></w:tr></w:tbl>` +
      p("Table 2: Costs") +
      "<w:p/>" +
      p("Annexure A");
    const result = await detectStructure(await makeDocx(body, STYLES), "a.docx");
    expect(result.paragraphs.map((d) => [d.role, d.level])).toEqual([
      ["title", null], ["heading", 1], ["heading", 2], ["body", null], ["figure-caption", null],
      ["list-item", null], ["table", null], ["table-caption", null], ["body", null], ["heading", 1],
    ]);
    expect(result.paragraphs.map((d) => d.paragraph.index)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(result.paragraphs.every((d) => d.manualRole === null)).toBe(true);
    expect(result.lowConfidenceCount).toBe(result.paragraphs.filter((d) => d.confidence < LOW_CONFIDENCE_THRESHOLD).length);
  });

  it("counts low-confidence paragraphs for manual tagging", async () => {
    const body = p("Normal text paragraph number one is here and is long enough to be body text.") +
      p("Notes", "", "<w:b/>") + p("Table 7 shows that costs went up in the last year.");
    const result = await detectStructure(await makeDocx(body, STYLES), "a.docx");
    expect(result.paragraphs[1]?.role).toBe("unknown");
    expect(result.paragraphs[2]?.confidence).toBeLessThan(LOW_CONFIDENCE_THRESHOLD);
    expect(result.lowConfidenceCount).toBe(2);
  });

  it("finds the body size by characters, so a few big lines do not win", async () => {
    const body = p("Overview", "", '<w:b/><w:sz w:val="32"/>') +
      p("Body text one is here and long enough.", "", '<w:sz w:val="24"/>') +
      p("Body text two is here and long enough.", "", '<w:sz w:val="24"/>');
    const summary = await detectStructureStream(await makeDocx(body, STYLES), "a.docx", () => undefined);
    expect(summary.bodySizeHalfPoints).toBe(24);
    expect(summary.paragraphCount).toBe(3);
  });

  it("uses the default size when the document has no styles.xml", async () => {
    const result = await detectStructure(await makeDocx(p("Hello there my friend")), "a.docx");
    expect(result.paragraphs[0]?.role).toBe("body");
  });

  it("a heading style also works without outline info and with another case", async () => {
    const result = await detectStructure(await makeDocx(p("Intro", '<w:pStyle w:val="Heading1"/>'), STYLES), "a.docx");
    expect(result.paragraphs[0]).toMatchObject({ role: "heading", level: 1 });
  });

  it("streams results one at a time and reports progress in order", async () => {
    const seen: DetectedParagraph[] = [];
    const progress: ProgressInfo[] = [];
    await detectStructureStream(await makeDocx(p("A b c") + p("D e f")), "a.docx", (d) => seen.push(d), { onProgress: (i) => progress.push(i) });
    expect(seen.length).toBe(2);
    expect(progress[0]?.fraction).toBe(0);
    expect(progress[progress.length - 1]?.fraction).toBe(1);
    expect(progress.every((i) => i.phase === "detecting")).toBe(true);
    for (let i = 1; i < progress.length; i += 1) expect(progress[i]?.fraction ?? 0).toBeGreaterThanOrEqual(progress[i - 1]?.fraction ?? 0);
  });

  it("does not change the source blob", async () => {
    const blob = await makeDocx(p("Hello"), STYLES);
    const before = new Uint8Array(await blob.arrayBuffer());
    await detectStructure(blob, "a.docx");
    const after = new Uint8Array(await blob.arrayBuffer());
    expect(after).toEqual(before);
  });
});

describe("detectStructure errors", () => {
  it("wrong extension -> NOT_DOCX", async () => {
    await expect(detectStructure(await makeDocx(p("x")), "a.txt")).rejects.toMatchObject({ code: "NOT_DOCX" });
  });
  it("DOCTYPE bomb in document.xml -> XML_ENTITY_EXPANSION", async () => {
    const blob = await makeDocx(p("x"), undefined, '<!DOCTYPE d [<!ENTITY a "aaaa"><!ENTITY b "&a;&a;&a;&a;">]>');
    await expect(detectStructure(blob, "a.docx")).rejects.toMatchObject({ code: "XML_ENTITY_EXPANSION" });
  });
  it("damaged XML -> ZIP_READ_FAILED with the part name", async () => {
    const { output, getBlob } = createBlobZipOutput();
    await output.addText("[Content_Types].xml", "<Types/>");
    await output.addText("word/document.xml", `<w:document ${NS}><w:body><w:p>`);
    await output.finish();
    await expect(detectStructure(await getBlob(), "a.docx")).rejects.toMatchObject({ code: "ZIP_READ_FAILED" });
  });
  it("already cancelled job -> CANCELLED", async () => {
    const ac = new AbortController();
    ac.abort();
    const control = createJobControl(60_000, ac.signal);
    try {
      await expect(detectStructure(await makeDocx(p("x")), "a.docx", { control })).rejects.toMatchObject({ code: "CANCELLED" });
    } finally {
      control.dispose();
    }
  });
  it("XML part over the byte cap -> an error, not a crash", async () => {
    const limits = { maxEntries: 100, maxTotalBytes: 1e9, maxXmlBytes: 400, timeoutMs: 60_000 };
    const blob = await makeDocx(p("x").repeat(50), STYLES);
    await expect(detectStructure(blob, "a.docx", { limits })).rejects.toMatchObject({ code: expect.stringMatching(/XML_TOO_LARGE|FILE_TOO_LARGE/) });
  });
});

describe("detectStructure on a real-style file made by the docx library", () => {
  it("headings, lists, tables and captions", async () => {
    const doc = new Document({
      numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: "bullet", text: "•", alignment: "left" }] }] },
      sections: [{
        children: [
          new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("Project Plan")] }),
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Introduction")] }),
          new Paragraph({ children: [new TextRun("Some body text that explains the project in a few words and a bit more.")] }),
          new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Scope")] }),
          new Paragraph({ numbering: { reference: "bul", level: 0 }, children: [new TextRun("First point")] }),
          new Paragraph({ children: [new TextRun("Table 1: Budget")] }),
          new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("Cost")] })] })] }),
          new Paragraph({ children: [new TextRun({ text: "2.3 Risks", bold: true })] }),
        ],
      }],
    });
    const buf = await Packer.toBuffer(doc);
    const result = await detectStructure(new Blob([new Uint8Array(buf)]), "real.docx");
    const roles = result.paragraphs.map((d) => d.role);
    expect(roles).toContain("title");
    expect(result.paragraphs.find((d) => d.paragraph.text === "Introduction")).toMatchObject({ role: "heading", level: 1 });
    expect(result.paragraphs.find((d) => d.paragraph.text === "Scope")).toMatchObject({ role: "heading", level: 2 });
    expect(result.paragraphs.find((d) => d.paragraph.text === "First point")?.role).toBe("list-item");
    expect(result.paragraphs.find((d) => d.paragraph.text === "Table 1: Budget")?.role).toBe("table-caption");
    expect(result.paragraphs.find((d) => d.paragraph.text === "Cost")?.role).toBe("table");
    expect(result.paragraphs.find((d) => d.paragraph.text === "2.3 Risks")).toMatchObject({ role: "heading", level: 2 });
  });
});

describe("detectStructure on a big document", () => {
  it("60000 paragraphs stream fast and in order", async () => {
    const parts: string[] = [];
    for (let i = 0; i < 60000; i += 1) {
      parts.push(i % 100 === 0 ? p(`${i / 100 + 1} Heading number ${i}`, '<w:pStyle w:val="Heading1"/>') : p("Body text of the paragraph with some ordinary words."));
    }
    const blob = await makeDocx(parts.join(""), STYLES);
    let count = 0;
    let headings = 0;
    let lastIndex = -1;
    const t0 = Date.now();
    await detectStructureStream(blob, "big.docx", (d) => {
      count += 1;
      if (d.role === "heading") headings += 1;
      expect(d.paragraph.index).toBe(lastIndex + 1);
      lastIndex = d.paragraph.index;
    });
    expect(count).toBe(60000);
    expect(headings).toBe(600);
    expect(Date.now() - t0).toBeLessThan(30000);
  }, 60000);
});
