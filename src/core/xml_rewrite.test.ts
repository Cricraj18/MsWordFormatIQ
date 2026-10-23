import { describe, expect, it } from "vitest";
import { FormatIQError } from "../types";
import {
  closeTag,
  createXmlRewriter,
  escapeXmlAttr,
  escapeXmlText,
  getAttr,
  openTag,
  rewriteXmlString,
  withAttr,
  XmlSerializer,
} from "./xml_rewrite";
import type { RawOpen, RawToken } from "./xml_rewrite";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const same = (xml: string): string => rewriteXmlString(xml, "test.xml");

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (error instanceof FormatIQError) return error.code;
    return "OTHER";
  }
  return "NONE";
}

describe("escape helpers", () => {
  it("escapes text", () => {
    expect(escapeXmlText('a & b < c > d "q"')).toBe('a &amp; b &lt; c &gt; d "q"');
  });
  it("escapes attributes incl. quotes, tabs and line breaks", () => {
    expect(escapeXmlAttr('a"b\tc\nd\re&<>')).toBe("a&quot;b&#9;c&#10;d&#13;e&amp;&lt;&gt;");
  });
});

describe("tag helpers", () => {
  const tag = openTag("w:pStyle", [["w:val", "Normal"]]);
  it("getAttr finds and misses", () => {
    expect(getAttr(tag, "w:val")).toBe("Normal");
    expect(getAttr(tag, "val")).toBeUndefined();
  });
  it("withAttr replaces in place and appends new, without changing the original", () => {
    const changed = withAttr(tag, "w:val", "Heading1");
    expect(getAttr(changed, "w:val")).toBe("Heading1");
    expect(getAttr(tag, "w:val")).toBe("Normal");
    const added = withAttr(tag, "w:x", "1");
    expect(added.attrs.map((a) => a[0])).toEqual(["w:val", "w:x"]);
  });
});

describe("XmlSerializer", () => {
  it("writes <a/> for an immediately closed element and <a></a> never", () => {
    const out: string[] = [];
    const s = new XmlSerializer((c) => out.push(c));
    s.push(openTag("a"));
    s.push(closeTag("a"));
    s.flush();
    expect(out.join("")).toBe("<a/>");
  });
  it("flushes a held open tag when something else follows", () => {
    const out: string[] = [];
    const s = new XmlSerializer((c) => out.push(c));
    s.push(openTag("a"));
    s.push({ kind: "text", text: "x" });
    s.push(closeTag("a"));
    s.flush();
    expect(out.join("")).toBe("<a>x</a>");
  });
  it("flush writes a dangling open tag", () => {
    const out: string[] = [];
    const s = new XmlSerializer((c) => out.push(c));
    s.push(openTag("a"));
    s.flush();
    expect(out.join("")).toBe("<a>");
  });
  it("skips empty text", () => {
    const out: string[] = [];
    const s = new XmlSerializer((c) => out.push(c));
    s.push({ kind: "text", text: "" });
    expect(out).toEqual([]);
  });
});

describe("round trip (no transform)", () => {
  it("keeps prefixes, attribute order and namespace declarations", () => {
    const xml = `<w:document ${NS} xmlns:r="urn:r"><w:body><w:p w:rsidR="00A1" w:rsidP="00B2"><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>`;
    expect(same(xml)).toBe(xml);
  });
  it("keeps the XML declaration", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}/>`;
    const out = same(xml);
    expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
    expect(out.endsWith(`<w:document ${NS}/>`)).toBe(true);
  });
  it("turns an empty pair into a self-closing tag", () => {
    expect(same(`<a><b></b><c/></a>`)).toBe("<a><b/><c/></a>");
  });
  it("keeps whitespace text exactly (xml:space preserve)", () => {
    const xml = `<w:t xml:space="preserve">  two  spaces </w:t>`;
    expect(same(xml)).toBe(xml);
  });
  it("re-escapes the five entities and numeric references", () => {
    expect(same(`<a b="x&amp;y&quot;z">1 &lt; 2 &amp; 3 &gt; 0 &#65;&#x42; &apos;</a>`)).toBe(
      `<a b="x&amp;y&quot;z">1 &lt; 2 &amp; 3 &gt; 0 AB '</a>`,
    );
  });
  it("keeps tab and line break characters inside attributes", () => {
    const out = same(`<a b="x&#9;y&#10;z"/>`);
    expect(out).toBe(`<a b="x&#9;y&#10;z"/>`);
  });
  it("keeps comments and processing instructions", () => {
    const xml = `<a><!-- note --><?target some data?><b/></a>`;
    expect(same(xml)).toBe(xml);
  });
  it("turns CDATA into escaped text with the same meaning", () => {
    expect(same(`<a><![CDATA[x < y & z]]></a>`)).toBe("<a>x &lt; y &amp; z</a>");
  });
  it("keeps accents, Japanese and emoji", () => {
    const xml = `<a>café 日本語 😀</a>`;
    expect(same(xml)).toBe(xml);
  });
  it("same result when fed one character at a time", () => {
    const xml = `<?xml version="1.0"?><w:p ${NS} w:a="1"><w:t>café 日本 😀 &amp; x</w:t><!--c--><w:br/></w:p>`;
    const whole = same(xml);
    const pieces: string[] = [];
    const rewriter = createXmlRewriter({ partName: "t.xml", onOutput: (c) => pieces.push(c) });
    for (const ch of Array.from(xml)) rewriter.write(ch);
    rewriter.end();
    expect(pieces.join("")).toBe(whole);
  });
  it("output parses again to the same output (stable)", () => {
    const xml = `<a x="1&amp;2"><b>t &lt; u</b><c/></a>`;
    const once = same(xml);
    expect(same(once)).toBe(once);
  });
});

describe("transform", () => {
  it("renames a style id", () => {
    const xml = `<w:p ${NS}><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const out = rewriteXmlString(xml, "t.xml", (token, emit) => {
      if (token.kind === "open" && token.name === "w:pStyle") emit(withAttr(token, "w:val", "Body"));
      else emit(token);
    });
    expect(out).toContain('<w:pStyle w:val="Body"/>');
    expect(out).not.toContain("Normal");
  });
  it("inserts a child into an empty element", () => {
    const out = rewriteXmlString(`<w:p/>`, "t.xml", (token, emit) => {
      emit(token);
      if (token.kind === "open" && token.name === "w:p") {
        emit(openTag("w:pPr"));
        emit(openTag("w:pStyle", [["w:val", "H1"]]));
        emit(closeTag("w:pStyle"));
        emit(closeTag("w:pPr"));
      }
    });
    expect(out).toBe('<w:p><w:pPr><w:pStyle w:val="H1"/></w:pPr></w:p>');
  });
  it("drops an element and its content when the transform tracks depth", () => {
    let skip = 0;
    const out = rewriteXmlString(`<a><b><c>gone</c></b><d>kept</d></a>`, "t.xml", (token, emit) => {
      if (token.kind === "open" && token.name === "b") skip += 1;
      if (skip === 0) emit(token);
      if (token.kind === "close" && token.name === "b") skip -= 1;
    });
    expect(out).toBe("<a><d>kept</d></a>");
  });
  it("can emit more than one token and raw xml", () => {
    const out = rewriteXmlString(`<a/>`, "t.xml", (token, emit) => {
      emit(token);
      if (token.kind === "open") emit({ kind: "raw", xml: "<!--added-->" });
    });
    expect(out).toBe("<a><!--added--></a>");
  });
  it("escapes new text and attribute values the transform adds", () => {
    const out = rewriteXmlString(`<a/>`, "t.xml", (token, emit) => {
      if (token.kind === "open") {
        emit(openTag("a", [["q", 'a"b&c']]));
        emit({ kind: "text", text: "1 < 2" });
      } else emit(token);
    });
    expect(out).toBe('<a q="a&quot;b&amp;c">1 &lt; 2</a>');
  });
  it("errors thrown by the transform come out unchanged", () => {
    const boom = new FormatIQError("CANCELLED", "stop");
    let caught: unknown;
    try {
      rewriteXmlString(`<a><b/></a>`, "t.xml", (token, emit) => {
        if (token.kind === "open" && token.name === "b") throw boom;
        emit(token);
      });
    } catch (error) {
      caught = error;
    }
    expect(caught).toBe(boom);
  });
  it("a transform that emits nothing gives an empty result", () => {
    expect(rewriteXmlString(`<a>x</a>`, "t.xml", () => undefined)).toBe("");
  });
});

describe("safety and errors", () => {
  it("refuses a DOCTYPE with an entity bomb", () => {
    const bomb = `<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;">]><r>&b;</r>`;
    expect(codeOf(() => same(bomb))).toBe("XML_ENTITY_EXPANSION");
  });
  it("refuses a plain DOCTYPE", () => {
    expect(codeOf(() => same(`<!DOCTYPE html><a/>`))).toBe("XML_ENTITY_EXPANSION");
  });
  it("damaged XML gives ZIP_READ_FAILED naming the part", () => {
    let detail: string | undefined;
    try {
      rewriteXmlString(`<a><b></a>`, "word/document.xml");
    } catch (error) {
      if (error instanceof FormatIQError) {
        expect(error.code).toBe("ZIP_READ_FAILED");
        detail = error.detail;
      }
    }
    expect(detail?.startsWith("word/document.xml:")).toBe(true);
  });
  it("unclosed document gives ZIP_READ_FAILED at end()", () => {
    expect(codeOf(() => same(`<a><b>`))).toBe("ZIP_READ_FAILED");
  });
  it("undefined entity is not expanded and gives an error", () => {
    expect(codeOf(() => same(`<a>&custom;</a>`))).toBe("ZIP_READ_FAILED");
  });
});

describe("streaming size", () => {
  it("rewrites 100000 paragraphs in pieces, in order, and changes one style", () => {
    const parts: string[] = [`<w:body ${NS}>`];
    for (let i = 0; i < 100000; i += 1) {
      parts.push(`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Line ${i}</w:t></w:r></w:p>`);
    }
    parts.push("</w:body>");
    const xml = parts.join("");
    let outputLength = 0;
    let changed = 0;
    let firstChunk = "";
    let lastChunk = "";
    const rewriter = createXmlRewriter({
      partName: "big.xml",
      onOutput: (chunk) => {
        outputLength += chunk.length;
        if (firstChunk === "") firstChunk = chunk;
        lastChunk = chunk;
      },
      transform: (token: RawToken, emit) => {
        if (token.kind === "open" && token.name === "w:pStyle") {
          changed += 1;
          const tag: RawOpen = withAttr(token, "w:val", "Body");
          emit(tag);
        } else emit(token);
      },
    });
    for (let i = 0; i < xml.length; i += 65536) rewriter.write(xml.slice(i, i + 65536));
    rewriter.end();
    expect(changed).toBe(100000);
    expect(firstChunk.startsWith("<w:body")).toBe(true);
    expect(lastChunk).toBe("</w:body>");
    expect(outputLength).toBe(xml.length - 100000 * ("Normal".length - "Body".length));
  });
});
