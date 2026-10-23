import { describe, expect, it } from "vitest";
import { FormatIQError } from "../types";
import {
  MAX_PAGE_TWIPS,
  MIN_PAGE_TWIPS,
  PAGE_SIZES,
  applyPlanToSectPr,
  checkMargins,
  createDefaultPageSetup,
  createPageSetupTransform,
  createSectionTransform,
  factsFromSectPr,
  inchesToTwips,
  mmToTwips,
  planSection,
  portraitSize,
  scanSections,
} from "./page_setup";
import type { MarginSet, PageSetupOptions, SectionFacts } from "./page_setup";
import { createXmlRewriter, rewriteXmlString } from "./xml_rewrite";
import type { RawToken } from "./xml_rewrite";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const doc = (body: string): string => `<w:document ${NS}><w:body>${body}</w:body></w:document>`;
const A4 = `<w:pgSz w:w="11906" w:h="16838" w:code="9"/>`;
const MAR = `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>`;
const sect = (inner: string): string => `<w:sectPr>${inner}</w:sectPr>`;
const para = (inner: string): string => `<w:p><w:pPr>${inner}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;

const TEMPLATE: SectionFacts = {
  index: 0,
  widthTwips: 11906,
  heightTwips: 16838,
  landscape: false,
  margins: { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 567, footer: 567, gutter: 0 },
  breakType: undefined,
};
const SOURCE: SectionFacts = { index: 0, widthTwips: 12240, heightTwips: 15840, landscape: false, margins: {}, breakType: undefined };
const opts = (over: Partial<PageSetupOptions>): PageSetupOptions => ({ ...createDefaultPageSetup(), ...over });

function run(xml: string, options: PageSetupOptions, template: SectionFacts | undefined = TEMPLATE) {
  const t = createPageSetupTransform(options, template);
  return { out: rewriteXmlString(xml, "document.xml", t.transform), stats: t.stats };
}

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    return error instanceof FormatIQError ? error.code : "OTHER";
  }
  return "NONE";
}

describe("units and presets", () => {
  it("converts millimetres and inches", () => {
    expect(mmToTwips(25.4)).toBe(1440);
    expect(mmToTwips(210)).toBe(11906);
    expect(mmToTwips(297)).toBe(16838);
    expect(inchesToTwips(8.5)).toBe(12240);
    expect(inchesToTwips(11)).toBe(15840);
  });
  it("preset table is portrait and matches Word", () => {
    for (const size of Object.values(PAGE_SIZES)) expect(size.width).toBeLessThan(size.height);
    expect(PAGE_SIZES.a4).toEqual({ width: 11906, height: 16838 });
    expect(PAGE_SIZES.letter).toEqual({ width: 12240, height: 15840 });
    expect(PAGE_SIZES.a3).toEqual({ width: 16838, height: 23811 });
  });
  it("default options use the template for everything", () => {
    expect(createDefaultPageSetup()).toEqual({ preset: "template", orientation: "template", margins: "template", sections: "all" });
  });
});

describe("reading sections", () => {
  it("finds sections in paragraphs and at the end of the body, in order", () => {
    const xml = doc(
      para(sect(`<w:type w:val="nextPage"/>${A4}${MAR}`)) +
        para(sect(`<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>`)) +
        sect(`<w:pgSz w:w="12240" w:h="15840"/>`),
    );
    const list = scanSections(xml);
    expect(list.map((s) => s.index)).toEqual([0, 1, 2]);
    expect(list[0]).toMatchObject({ widthTwips: 11906, heightTwips: 16838, landscape: false, breakType: "nextPage" });
    expect(list[0]?.margins).toEqual({ top: 1440, right: 1440, bottom: 1440, left: 1440, header: 708, footer: 708, gutter: 0 });
    expect(list[1]).toMatchObject({ widthTwips: 16838, heightTwips: 11906, landscape: true });
    expect(list[2]?.landscape).toBe(false);
  });
  it("landscape without orient is read from width larger than height", () => {
    expect(scanSections(doc(sect(`<w:pgSz w:w="16838" w:h="11906"/>`)))[0]?.landscape).toBe(true);
  });
  it("a section without pgSz or pgMar has undefined sizes and empty margins", () => {
    const s = scanSections(doc(sect(`<w:cols w:space="708"/>`)))[0];
    expect(s?.widthTwips).toBeUndefined();
    expect(s?.landscape).toBe(false);
    expect(s?.margins).toEqual({});
  });
  it("a sectPrChange record is part of its section and not a section of its own", () => {
    const xml = doc(sect(`${A4}<w:sectPrChange w:id="1"><w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr></w:sectPrChange>`));
    const list = scanSections(xml);
    expect(list).toHaveLength(1);
    expect(list[0]?.widthTwips).toBe(11906);
  });
  it("a document without sections gives an empty list", () => {
    expect(scanSections(doc("<w:p/>"))).toEqual([]);
  });
});

describe("portraitSize and checkMargins", () => {
  it("template size, A4 fallback, presets and custom", () => {
    expect(portraitSize(opts({}), TEMPLATE)).toEqual({ short: 11906, long: 16838 });
    expect(portraitSize(opts({}), undefined)).toEqual({ short: 11906, long: 16838 });
    expect(portraitSize(opts({}), { ...TEMPLATE, widthTwips: 16838, heightTwips: 11906 })).toEqual({ short: 11906, long: 16838 });
    expect(portraitSize(opts({ preset: "letter" }), TEMPLATE)).toEqual({ short: 12240, long: 15840 });
    expect(portraitSize(opts({ preset: "custom", customWidthTwips: 20000, customHeightTwips: 10000 }), TEMPLATE)).toEqual({ short: 10000, long: 20000 });
  });
  it("custom size needs both sides inside the allowed range", () => {
    expect(codeOf(() => portraitSize(opts({ preset: "custom" }), TEMPLATE))).toBe("INVALID_OPTIONS");
    expect(codeOf(() => portraitSize(opts({ preset: "custom", customWidthTwips: MIN_PAGE_TWIPS - 1, customHeightTwips: 5000 }), TEMPLATE))).toBe("INVALID_OPTIONS");
    expect(codeOf(() => portraitSize(opts({ preset: "custom", customWidthTwips: 5000, customHeightTwips: MAX_PAGE_TWIPS + 1 }), TEMPLATE))).toBe("INVALID_OPTIONS");
    expect(codeOf(() => portraitSize(opts({ preset: "custom", customWidthTwips: 5000.5, customHeightTwips: 6000 }), TEMPLATE))).toBe("INVALID_OPTIONS");
    expect(codeOf(() => portraitSize(opts({ preset: "custom", customWidthTwips: MIN_PAGE_TWIPS, customHeightTwips: MAX_PAGE_TWIPS }), TEMPLATE))).toBe("NONE");
  });
  it("margins that fit give no problems", () => {
    expect(checkMargins(TEMPLATE.margins, 11906, 16838)).toEqual([]);
    expect(checkMargins({}, 11906, 16838)).toEqual([]);
  });
  it("negative, fractional, and too wide margins are reported", () => {
    expect(checkMargins({ top: -1 }, 11906, 16838)).toHaveLength(1);
    expect(checkMargins({ left: 10.5 }, 11906, 16838)).toHaveLength(1);
    expect(checkMargins({ left: 6000, right: 6000 }, 11906, 16838).join(" ")).toContain("left and right");
    expect(checkMargins({ top: 9000, bottom: 9000 }, 11906, 16838).join(" ")).toContain("top and bottom");
    expect(checkMargins({ left: 5000, right: 5000, gutter: 1500 }, 11906, 16838).join(" ")).toContain("left and right");
  });
});

describe("planSection", () => {
  it("template size and margins by default", () => {
    expect(planSection(opts({}), TEMPLATE, SOURCE)).toEqual({ width: 11906, height: 16838, landscape: false, margins: TEMPLATE.margins });
  });
  it("preset sets the size, landscape swaps it", () => {
    expect(planSection(opts({ preset: "a3" }), TEMPLATE, SOURCE)).toMatchObject({ width: 16838, height: 23811, landscape: false });
    expect(planSection(opts({ preset: "a3", orientation: "landscape" }), TEMPLATE, SOURCE)).toMatchObject({ width: 23811, height: 16838, landscape: true });
  });
  it("portrait forces portrait even for a landscape source section", () => {
    const wide: SectionFacts = { ...SOURCE, landscape: true, widthTwips: 15840, heightTwips: 12240 };
    expect(planSection(opts({ orientation: "portrait" }), TEMPLATE, wide)).toMatchObject({ landscape: false, width: 11906 });
  });
  it("orientation template keeps a landscape source section landscape (swapped chosen size)", () => {
    const wide: SectionFacts = { ...SOURCE, landscape: true };
    expect(planSection(opts({ preset: "letter" }), TEMPLATE, wide)).toMatchObject({ landscape: true, width: 15840, height: 12240 });
  });
  it("orientation template follows a landscape template", () => {
    const wideTemplate: SectionFacts = { ...TEMPLATE, landscape: true, widthTwips: 16838, heightTwips: 11906 };
    expect(planSection(opts({}), wideTemplate, SOURCE)).toMatchObject({ landscape: true, width: 16838, height: 11906 });
  });
  it("custom margins replace the template margins", () => {
    const margins: MarginSet = { top: 1000, right: 900, bottom: 1000, left: 900, header: 400, footer: 400, gutter: 100 };
    expect(planSection(opts({ margins }), TEMPLATE, SOURCE)?.margins).toEqual(margins);
  });
  it("a template without margins leaves margins empty", () => {
    expect(planSection(opts({}), { ...TEMPLATE, margins: {} }, SOURCE)?.margins).toEqual({});
  });
  it("only selected sections get a plan", () => {
    expect(planSection(opts({ sections: [1] }), TEMPLATE, { ...SOURCE, index: 0 })).toBeUndefined();
    expect(planSection(opts({ sections: [1] }), TEMPLATE, { ...SOURCE, index: 1 })).toBeDefined();
    expect(planSection(opts({ sections: [] }), TEMPLATE, SOURCE)).toBeUndefined();
  });
  it("bad margins for this page size throw INVALID_OPTIONS and name the section", () => {
    const margins: MarginSet = { top: 1, right: 9000, bottom: 1, left: 9000, header: 1, footer: 1, gutter: 0 };
    let detail: string | undefined;
    try {
      planSection(opts({ margins }), TEMPLATE, { ...SOURCE, index: 2 });
    } catch (error) {
      if (error instanceof FormatIQError) {
        expect(error.code).toBe("INVALID_OPTIONS");
        detail = error.detail;
      }
    }
    expect(detail).toBe("section 3");
  });
});

describe("applying to a document", () => {
  it("sets size, margins and removes the paper code (A4 to Letter)", () => {
    const { out } = run(doc(sect(`${A4}${MAR}`)), opts({ preset: "letter", margins: { top: 1000, right: 1100, bottom: 1200, left: 1300, header: 500, footer: 600, gutter: 0 } }));
    expect(out).toContain(`<w:pgSz w:w="12240" w:h="15840"/>`);
    expect(out).toContain(`<w:pgMar w:top="1000" w:right="1100" w:bottom="1200" w:left="1300" w:header="500" w:footer="600" w:gutter="0"/>`);
    expect(out).not.toContain("w:code");
  });
  it("writes orient=landscape and swaps the size; portrait removes orient", () => {
    const land = run(doc(sect(`${A4}${MAR}`)), opts({ orientation: "landscape" })).out;
    expect(land).toContain(`<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>`);
    const back = run(land, opts({ orientation: "portrait" })).out;
    expect(back).toContain(`<w:pgSz w:w="11906" w:h="16838"/>`);
    expect(back).not.toContain("orient");
  });
  it("keeps every other child of the section untouched", () => {
    const others = `<w:headerReference w:type="default" r:id="rId9"/><w:footerReference w:type="default" r:id="rId10"/><w:type w:val="continuous"/>`;
    const tail = `<w:pgNumType w:start="1"/><w:cols w:num="2" w:space="708"/><w:titlePg/><w:docGrid w:linePitch="360"/>`;
    const { out } = run(doc(sect(`${others}${A4}${MAR}${tail}`)), opts({ preset: "a3" }));
    expect(out).toContain(`${others}<w:pgSz`);
    expect(out).toContain(`${tail}</w:sectPr>`);
  });
  it("creates missing pgSz and pgMar after headers, footnote settings and type", () => {
    const before = `<w:headerReference w:type="default" r:id="rId9"/><w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:type w:val="nextPage"/>`;
    const after = `<w:cols w:space="708"/>`;
    const { out } = run(doc(sect(`${before}${after}`)), opts({}));
    expect(out).toContain(`${before}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>${after}`);
  });
  it("creates pgSz in an empty sectPr", () => {
    const { out } = run(doc(`<w:sectPr/>`), opts({}, ), { ...TEMPLATE, margins: {} });
    expect(out).toContain(`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`);
  });
  it("keeps the old pgMar when there are no margins to write", () => {
    const { out } = run(doc(sect(`${A4}${MAR}`)), opts({}), { ...TEMPLATE, margins: {} });
    expect(out).toContain(MAR);
  });
  it("adds pgMar after an existing pgSz when only the margins were missing", () => {
    const { out } = run(doc(sect(`${A4}<w:cols w:space="708"/>`)), opts({}));
    expect(out).toContain(`<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134"`);
    expect(out.indexOf("<w:pgMar")).toBeLessThan(out.indexOf("<w:cols"));
  });
  it("changes only the selected sections; others are copied exactly", () => {
    const first = para(sect(`${A4}${MAR}`));
    const second = sect(`<w:pgSz w:w="12240" w:h="15840"/>${MAR}`);
    const { out, stats } = run(doc(first + second), opts({ preset: "a3", sections: [1] }));
    expect(out).toContain(first);
    expect(out).toContain(`<w:pgSz w:w="16838" w:h="23811"/>`);
    expect(stats).toEqual({ sections: 2, changed: 1 });
  });
  it("source landscape section stays landscape with the new paper", () => {
    const xml = doc(para(sect(`${A4}${MAR}`)) + para(sect(`<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>${MAR}`)) + sect(`${A4}${MAR}`));
    const { out } = run(xml, opts({ preset: "letter" }));
    const list = scanSections(out);
    expect(list.map((s) => [s.widthTwips, s.heightTwips, s.landscape])).toEqual([
      [12240, 15840, false],
      [15840, 12240, true],
      [12240, 15840, false],
    ]);
  });
  it("does not touch the record inside sectPrChange", () => {
    const rec = `<w:sectPrChange w:id="1"><w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr></w:sectPrChange>`;
    const { out } = run(doc(sect(`${A4}${MAR}${rec}`)), opts({ preset: "a3" }));
    expect(out).toContain(rec);
    expect(scanSections(out)[0]?.widthTwips).toBe(16838);
  });
  it("uses the document's own prefix for new attributes", () => {
    const xml = `<x:document xmlns:x="urn:w"><x:body><x:sectPr><x:cols x:space="1"/></x:sectPr></x:body></x:document>`;
    const { out } = run(xml, opts({ orientation: "landscape" }));
    expect(out).toContain(`<x:pgSz x:w="16838" x:h="11906" x:orient="landscape"/>`);
    expect(out).toContain(`<x:pgMar x:top="1134"`);
  });
  it("text and other content are not changed", () => {
    const body = `<w:p><w:r><w:t>Hello &amp; goodbye</w:t></w:r></w:p>` + para(sect(`${A4}${MAR}`)) + `<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    const { out } = run(doc(body), opts({ preset: "letter" }));
    expect(out).toContain(`<w:p><w:r><w:t>Hello &amp; goodbye</w:t></w:r></w:p>`);
    expect(out).toContain(`<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`);
  });
  it("is stable when run twice with the same options", () => {
    const options = opts({ preset: "a3", orientation: "landscape" });
    const once = run(doc(para(sect(`${A4}${MAR}`)) + sect(`${A4}${MAR}`)), options).out;
    expect(run(once, options).out).toBe(once);
  });
  it("a document without sections is copied unchanged", () => {
    const xml = doc("<w:p/>");
    const { out, stats } = run(xml, opts({ preset: "a3" }));
    expect(out).toBe(xml);
    expect(stats).toEqual({ sections: 0, changed: 0 });
  });
  it("an invalid custom size stops at the first section", () => {
    expect(codeOf(() => run(doc(sect(A4)), opts({ preset: "custom" })))).toBe("INVALID_OPTIONS");
  });
});

describe("building blocks", () => {
  it("applyPlanToSectPr leaves a token list without an open tag alone", () => {
    const tokens: RawToken[] = [{ kind: "text", text: "x" }];
    expect(applyPlanToSectPr(tokens, { width: 1, height: 2, landscape: false, margins: {} })).toEqual(tokens);
  });
  it("factsFromSectPr reads one buffered section", () => {
    const seen: SectionFacts[] = [];
    const t = createSectionTransform((index, tokens) => {
      seen.push(factsFromSectPr(index, tokens));
      return tokens;
    });
    const rewriter = createXmlRewriter({ partName: "t.xml", onOutput: () => undefined, transform: t });
    rewriter.write(doc(sect(`<w:type w:val="evenPage"/>${A4}`)));
    rewriter.end();
    expect(seen).toHaveLength(1);
    expect(seen[0]?.breakType).toBe("evenPage");
  });
});

describe("streaming size", () => {
  it("changes 20000 sections fed in pieces", () => {
    const parts: string[] = [`<w:body ${NS}>`];
    for (let i = 0; i < 20000; i += 1) parts.push(para(sect(`${A4}${MAR}`)));
    parts.push("</w:body>");
    const xml = parts.join("");
    const t = createPageSetupTransform(opts({ preset: "letter" }), TEMPLATE);
    let count = 0;
    const rewriter = createXmlRewriter({ partName: "big.xml", onOutput: (c) => { count += c.length; }, transform: t.transform });
    for (let i = 0; i < xml.length; i += 65536) rewriter.write(xml.slice(i, i + 65536));
    rewriter.end();
    expect(t.stats).toEqual({ sections: 20000, changed: 20000 });
    expect(count).toBeGreaterThan(0);
  });
});
