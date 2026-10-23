import { beforeAll, describe, expect, it } from "vitest";
import { createXmlFeeder, localName, parseEntryXml } from "./xml";
import type { XmlHandlers } from "./xml";
import { FormatIQError } from "../types";
import { createBlobZipOutput, openDocx, setWorkersEnabled } from "./zipio";

beforeAll(() => setWorkersEnabled(false));

function collect(xml: string, pieceSize = xml.length): { opens: string[]; text: string; closes: string[]; attrs: Record<string, string>[] } {
  const out = { opens: [] as string[], text: "", closes: [] as string[], attrs: [] as Record<string, string>[] };
  const handlers: XmlHandlers = {
    open: (t) => { out.opens.push(t.name); out.attrs.push({ ...t.attrs }); },
    close: (n) => { out.closes.push(n); },
    text: (s) => { out.text += s; },
  };
  const feeder = createXmlFeeder(handlers, "test.xml");
  for (let i = 0; i < xml.length; i += pieceSize) feeder.write(xml.slice(i, i + pieceSize));
  feeder.end();
  return out;
}

describe("localName", () => {
  it("removes the prefix", () => {
    expect(localName("w:val")).toBe("val");
    expect(localName("plain")).toBe("plain");
  });
});

describe("createXmlFeeder", () => {
  it("reads tags, attributes and text, prefixes removed", () => {
    const r = collect('<w:p w:rsidR="1"><w:t>Hi</w:t><w:br/></w:p>');
    expect(r.opens).toEqual(["p", "t", "br"]);
    expect(r.closes).toEqual(["t", "br", "p"]);
    expect(r.attrs[0]).toEqual({ rsidR: "1" });
    expect(r.text).toBe("Hi");
  });

  it("gives the same result when fed one character at a time", () => {
    const xml = '<r a="1"><b>x&amp;y</b><![CDATA[<z>]]></r>';
    const whole = collect(xml);
    const tiny = collect(xml, 1);
    expect(tiny.opens).toEqual(whole.opens);
    expect(tiny.text).toBe(whole.text);
    expect(whole.text).toBe("x&y<z>");
  });

  it("decodes only built-in and numeric entities", () => {
    expect(collect("<r>&lt;&gt;&quot;&apos;&#65;&#x42;</r>").text).toBe("<>\"'AB");
  });

  it("rejects a DOCTYPE with XML_ENTITY_EXPANSION (billion laughs)", () => {
    const bomb = '<?xml version="1.0"?><!DOCTYPE l [<!ENTITY a "lol"><!ENTITY b "&a;&a;&a;&a;">]><r>&b;</r>';
    try {
      collect(bomb);
      throw new Error("should have thrown");
    } catch (err) {
      expect(err).toBeInstanceOf(FormatIQError);
      expect((err as FormatIQError).code).toBe("XML_ENTITY_EXPANSION");
    }
  });

  it("rejects an unknown entity as damaged XML", () => {
    expect(() => collect("<r>&foo;</r>")).toThrowError(FormatIQError);
  });

  it("rejects unclosed XML at the end", () => {
    expect(() => collect("<r><a>")).toThrowError(FormatIQError);
  });
});

async function zipWithXml(path: string, xml: string): Promise<{ entry: Awaited<ReturnType<typeof openDocx>>["entries"] extends Map<string, infer E> ? E : never; close: () => Promise<void> }> {
  const { output, getBlob } = createBlobZipOutput();
  await output.addText("[Content_Types].xml", "<Types/>");
  await output.addText("word/document.xml", "<d/>");
  await output.addText(path, xml);
  await output.finish();
  const docx = await openDocx(await getBlob(), "x.docx");
  const entry = docx.entries.get(path);
  if (entry === undefined) throw new Error("missing entry");
  return { entry, close: () => docx.close() };
}

describe("parseEntryXml", () => {
  it("streams a multi-byte text correctly", async () => {
    const { entry, close } = await zipWithXml("word/x.xml", "<r>héllo — 日本語 😀</r>");
    let text = "";
    await parseEntryXml(entry, { text: (t) => { text += t; } }, 1_000_000);
    await close();
    expect(text).toBe("héllo — 日本語 😀");
  });

  it("streams a big file (about 12 MB, 100000 paragraphs)", async () => {
    const body = "<p><t>Hello world text</t></p>".repeat(100_000);
    const { entry, close } = await zipWithXml("word/big.xml", `<root>${body}</root>`);
    let count = 0;
    await parseEntryXml(entry, { open: (t) => { if (t.name === "p") count += 1; } }, 50_000_000);
    await close();
    expect(count).toBe(100_000);
  });

  it("stops with FILE_TOO_LARGE when over the byte cap", async () => {
    const { entry, close } = await zipWithXml("word/x.xml", `<r>${"a".repeat(5000)}</r>`);
    await expect(parseEntryXml(entry, {}, 1000)).rejects.toMatchObject({ code: "FILE_TOO_LARGE" });
    await close();
  });

  it("passes XML_ENTITY_EXPANSION up through the zip stream", async () => {
    const { entry, close } = await zipWithXml("word/x.xml", '<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>');
    await expect(parseEntryXml(entry, {}, 1_000_000)).rejects.toMatchObject({ code: "XML_ENTITY_EXPANSION" });
    await close();
  });

  it("reports damaged XML as ZIP_READ_FAILED and names the part", async () => {
    const { entry, close } = await zipWithXml("word/x.xml", "<a><b></a>");
    await expect(parseEntryXml(entry, {}, 1_000_000)).rejects.toMatchObject({ code: "ZIP_READ_FAILED", detail: expect.stringContaining("word/x.xml") });
    await close();
  });
});
