import { describe, expect, it } from "vitest";
import {
  classifyParagraph, isShortLine, levelFromSizeRatio, matchCaption, matchLabeledHeading,
  matchNumberedHeading, pickBodySize, roleFromStyle,
} from "./structure_rules";
import type { DetectContext } from "./structure_rules";
import type { ParagraphFacts, StyleHint } from "./paragraphs";

const para = (over: Partial<ParagraphFacts>): ParagraphFacts => ({
  index: 0, text: "", styleId: null, inTable: false, hasImage: false, isListItem: false,
  outlineLevel: null, bold: null, sizeHalfPoints: null, ...over,
});
const styles = new Map<string, StyleHint>([
  ["Heading2", { name: "heading 2", outlineLevel: 2, bold: true, sizeHalfPoints: 28 }],
  ["Title", { name: "Title", outlineLevel: null, bold: null, sizeHalfPoints: 52 }],
  ["Cap", { name: "caption", outlineLevel: null, bold: null, sizeHalfPoints: null }],
  ["Odd", { name: "My Custom", outlineLevel: 3, bold: null, sizeHalfPoints: null }],
  ["BigBold", { name: "Big", outlineLevel: null, bold: true, sizeHalfPoints: 32 }],
  ["QuoteStyle", { name: "Quote", outlineLevel: null, bold: null, sizeHalfPoints: null }],
  ["Mono", { name: "Source Code", outlineLevel: null, bold: null, sizeHalfPoints: null }],
]);
const ctx = (over: Partial<DetectContext> = {}): DetectContext => ({
  styles, bodySizeHalfPoints: 22, defaultSizeHalfPoints: 22, nonEmptyBefore: 10, headingSeen: true, ...over,
});

describe("matchNumberedHeading", () => {
  it("matches 1 Introduction, 2.1 Scope, 3.2.1. Details with the right level", () => {
    expect(matchNumberedHeading("1 Introduction")?.level).toBe(1);
    expect(matchNumberedHeading("2.1 Scope")?.level).toBe(2);
    expect(matchNumberedHeading("3.2.1. Details")?.level).toBe(3);
    expect(matchNumberedHeading("1. Introduction")?.level).toBe(1);
  });
  it("dotted numbers are more sure than a single number", () => {
    expect(matchNumberedHeading("2.1 Scope")?.confidence).toBeGreaterThan(matchNumberedHeading("2 Scope")?.confidence ?? 1);
  });
  it("rejects sentences, lowercase starts, long text, years and plain numbers", () => {
    expect(matchNumberedHeading("3 apples were eaten")).toBeNull();
    expect(matchNumberedHeading("1. Buy milk and eggs from the shop today before noon please ok.")).toBeNull();
    expect(matchNumberedHeading("2024 Annual Report")).toBeNull();
    expect(matchNumberedHeading("12")).toBeNull();
    expect(matchNumberedHeading("5 This is a very long line that goes on and on and on and on forever")).toBeNull();
    expect(matchNumberedHeading("4 Results, discussed.")).toBeNull();
  });
  it("accepts non-Latin titles that have no letter case", () => {
    expect(matchNumberedHeading("1 はじめに")?.level).toBe(1);
  });
});

describe("matchLabeledHeading", () => {
  it("matches Chapter, Part, Section, Annexure, Appendix", () => {
    expect(matchLabeledHeading("Chapter 4")).toMatchObject({ level: 1, confidence: 0.9 });
    expect(matchLabeledHeading("CHAPTER 4: Methods")?.level).toBe(1);
    expect(matchLabeledHeading("Part II")?.level).toBe(1);
    expect(matchLabeledHeading("Section 3 Scope")?.level).toBe(2);
    expect(matchLabeledHeading("Annexure A")?.level).toBe(1);
    expect(matchLabeledHeading("Appendix B - Tables")?.level).toBe(1);
    expect(matchLabeledHeading("Chapter Three")?.level).toBe(1);
  });
  it("lowers confidence for sentences, rejects unrelated text", () => {
    expect(matchLabeledHeading("Section 3 applies to all members.")?.confidence).toBe(0.5);
    expect(matchLabeledHeading("Annexure A shows the data.")?.confidence).toBe(0.5);
    expect(matchLabeledHeading("Chapters are long")).toBeNull();
    expect(matchLabeledHeading("Appendix and more")).toBeNull();
    expect(matchLabeledHeading("Introduction")).toBeNull();
  });
});

describe("matchCaption", () => {
  it("figure and table captions", () => {
    expect(matchCaption("Figure 3: Flow of work")).toMatchObject({ role: "figure-caption", confidence: 0.92 });
    expect(matchCaption("Fig. 2.1 - Layout")?.role).toBe("figure-caption");
    expect(matchCaption("Table 7")?.role).toBe("table-caption");
    expect(matchCaption("TABLE 12. Costs")?.role).toBe("table-caption");
  });
  it("a sentence that mentions a table is weak, a long one is not a caption", () => {
    expect(matchCaption("Table 7 shows that costs went up")?.confidence).toBeLessThan(0.6);
    expect(matchCaption(`Figure 1: ${"x".repeat(300)}`)).toBeNull();
    expect(matchCaption("Tablet 7 is here")).toBeNull();
  });
});

describe("roleFromStyle", () => {
  it("heading n from name or id, title, quote, code", () => {
    expect(roleFromStyle("Heading3", undefined)).toEqual({ role: "heading", level: 3 });
    expect(roleFromStyle("X", { name: "heading 5", outlineLevel: null, bold: null, sizeHalfPoints: null })).toEqual({ role: "heading", level: 5 });
    expect(roleFromStyle("Title", undefined)?.role).toBe("title");
    expect(roleFromStyle("Quote", undefined)?.role).toBe("quote");
    expect(roleFromStyle("a", { name: "Source Code", outlineLevel: null, bold: null, sizeHalfPoints: null })?.role).toBe("code");
    expect(roleFromStyle("Normal", undefined)).toBeNull();
    expect(roleFromStyle(null, undefined)).toBeNull();
    expect(roleFromStyle("Heading10", undefined)).toBeNull();
  });
});

describe("small helpers", () => {
  it("isShortLine", () => {
    expect(isShortLine("Summary", 80)).toBe(true);
    expect(isShortLine("This is a sentence.", 80)).toBe(false);
    expect(isShortLine("", 80)).toBe(false);
    expect(isShortLine("x".repeat(90), 80)).toBe(false);
  });
  it("levelFromSizeRatio", () => {
    expect([2, 1.6, 1.4, 1.3, 1.2].map(levelFromSizeRatio)).toEqual([1, 1, 2, 2, 3]);
  });
  it("pickBodySize takes the heaviest, ties go to the smaller size, empty is null", () => {
    expect(pickBodySize(new Map([[22, 900], [28, 100]]))).toBe(22);
    expect(pickBodySize(new Map([[24, 50], [20, 50]]))).toBe(20);
    expect(pickBodySize(new Map())).toBeNull();
  });
});

describe("classifyParagraph", () => {
  it("table first", () => {
    expect(classifyParagraph(para({ text: "x", inTable: true, styleId: "Heading2" }), ctx()).role).toBe("table");
  });
  it("empty paragraph is body (picture alone 0.8, blank 1)", () => {
    expect(classifyParagraph(para({ text: "  ", hasImage: true }), ctx())).toMatchObject({ role: "body", confidence: 0.8 });
    expect(classifyParagraph(para({}), ctx())).toMatchObject({ role: "body", confidence: 1 });
  });
  it("Word styles: heading level, title, quote, code", () => {
    expect(classifyParagraph(para({ text: "Scope", styleId: "Heading2" }), ctx())).toEqual({ role: "heading", level: 2, confidence: 0.95 });
    expect(classifyParagraph(para({ text: "My Report", styleId: "Title" }), ctx()).role).toBe("title");
    expect(classifyParagraph(para({ text: "q", styleId: "QuoteStyle" }), ctx()).role).toBe("quote");
    expect(classifyParagraph(para({ text: "x = 1", styleId: "Mono" }), ctx()).role).toBe("code");
  });
  it("outline level on the paragraph or on a custom style gives a heading", () => {
    expect(classifyParagraph(para({ text: "Part", outlineLevel: 3 }), ctx())).toMatchObject({ role: "heading", level: 3, confidence: 0.9 });
    expect(classifyParagraph(para({ text: "Part", styleId: "Odd" }), ctx())).toMatchObject({ role: "heading", level: 3 });
  });
  it("captions by text and by Caption style", () => {
    expect(classifyParagraph(para({ text: "Figure 3: Map" }), ctx()).role).toBe("figure-caption");
    expect(classifyParagraph(para({ text: "Table 7: Costs" }), ctx()).role).toBe("table-caption");
    const r = classifyParagraph(para({ text: "Some words", styleId: "Cap" }), ctx());
    expect(r.role).toBe("figure-caption");
    expect(r.confidence).toBeLessThan(0.6);
  });
  it("text patterns make headings; bold raises the confidence", () => {
    expect(classifyParagraph(para({ text: "Chapter 4" }), ctx())).toMatchObject({ role: "heading", level: 1 });
    expect(classifyParagraph(para({ text: "2.1 Scope" }), ctx())).toMatchObject({ role: "heading", level: 2 });
    const plain = classifyParagraph(para({ text: "2 Scope" }), ctx()).confidence;
    const bold = classifyParagraph(para({ text: "2 Scope", bold: true }), ctx()).confidence;
    expect(bold).toBeGreaterThan(plain);
  });
  it("a numbered text line inside a real list stays a list item", () => {
    expect(classifyParagraph(para({ text: "1 Introduction", isListItem: true }), ctx()).role).toBe("list-item");
  });
  it("list numbering gives list-item", () => {
    expect(classifyParagraph(para({ text: "Buy milk", isListItem: true }), ctx())).toMatchObject({ role: "list-item", confidence: 0.9 });
  });
  it("big bold short line: title near the start, heading later", () => {
    const early = ctx({ nonEmptyBefore: 0, headingSeen: false });
    expect(classifyParagraph(para({ text: "Annual Report", bold: true, sizeHalfPoints: 48 }), early).role).toBe("title");
    expect(classifyParagraph(para({ text: "Annual Report", bold: true, sizeHalfPoints: 48 }), ctx())).toMatchObject({ role: "heading", level: 1 });
    expect(classifyParagraph(para({ text: "Overview", bold: true, sizeHalfPoints: 28 }), ctx())).toMatchObject({ role: "heading", level: 3, confidence: 0.65 });
  });
  it("style bold and size are used when the runs say nothing", () => {
    expect(classifyParagraph(para({ text: "Overview", styleId: "BigBold" }), ctx()).role).toBe("heading");
  });
  it("big but not bold is a weak guess; bold at body size is unknown", () => {
    const big = classifyParagraph(para({ text: "Overview", bold: false, sizeHalfPoints: 36 }), ctx());
    expect(big.role).toBe("heading");
    expect(big.confidence).toBeLessThan(0.6);
    const b = classifyParagraph(para({ text: "Note", bold: true }), ctx());
    expect(b).toMatchObject({ role: "unknown" });
    expect(b.confidence).toBeLessThan(0.6);
  });
  it("long or sentence-like big text is body", () => {
    expect(classifyParagraph(para({ text: "This is a normal long sentence in big letters.", bold: true, sizeHalfPoints: 36 }), ctx()).role).toBe("body");
    expect(classifyParagraph(para({ text: "Normal text here and there" }), ctx())).toMatchObject({ role: "body", confidence: 0.85 });
  });
});
