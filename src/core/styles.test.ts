import { describe, expect, it } from "vitest";
import type { DetectedParagraph, StructureRole, TemplateStyle } from "../types";
import {
  buildStyleMap,
  effectiveRole,
  headingStyleFor,
  missingStyleNames,
  normalizeStyleName,
  styleIdForParagraph,
  styleIdForRole,
} from "./styles";

const p = (id: string, name: string): TemplateStyle => ({ id, name, kind: "paragraph" });
const FULL: TemplateStyle[] = [
  p("Normal", "Normal"),
  p("Title", "Title"),
  p("Heading1", "heading 1"),
  p("Heading2", "heading 2"),
  p("Heading3", "heading 3"),
  p("Caption", "caption"),
  p("ListParagraph", "List Paragraph"),
  p("Quote", "Quote"),
  p("SourceCode", "Source Code"),
];

function detected(role: StructureRole, level: number | null, manual: StructureRole | null = null): DetectedParagraph {
  return {
    paragraph: { index: 0, text: "x", styleId: null, inTable: false, hasImage: false, isListItem: false },
    role,
    level,
    confidence: 0.9,
    manualRole: manual,
  };
}

describe("normalizeStyleName", () => {
  it("ignores case, spaces and dashes", () => {
    expect(normalizeStyleName("Heading 1")).toBe("heading1");
    expect(normalizeStyleName("heading-1")).toBe("heading1");
    expect(normalizeStyleName(" List  Paragraph ")).toBe("listparagraph");
  });
});

describe("buildStyleMap", () => {
  it("finds every style in a full template", () => {
    const m = buildStyleMap(FULL);
    expect(m.headings.slice(0, 3)).toEqual(["Heading1", "Heading2", "Heading3"]);
    expect(m.headings.slice(3).every((h) => h === null)).toBe(true);
    expect(m.headings).toHaveLength(9);
    expect(m.title).toBe("Title");
    expect(m.body).toBe("Normal");
    expect(m.listItem).toBe("ListParagraph");
    expect(m.figureCaption).toBe("Caption");
    expect(m.tableCaption).toBe("Caption");
    expect(m.quote).toBe("Quote");
    expect(m.code).toBe("SourceCode");
  });
  it("uses its own ids, whatever they are called", () => {
    const m = buildStyleMap([p("MyH1", "Heading 1"), p("Corpo", "Normal")]);
    expect(m.headings[0]).toBe("MyH1");
    expect(m.body).toBe("Corpo");
  });
  it("falls back to the id pattern when the name is localized", () => {
    const m = buildStyleMap([p("Heading1", "berschrift 1"), p("Normal", "Standard")]);
    expect(m.headings[0]).toBe("Heading1");
    expect(m.body).toBe("Normal");
  });
  it("prefers Figure Caption / Table Caption over Caption", () => {
    const m = buildStyleMap([p("Caption", "caption"), p("FigCap", "Figure Caption"), p("TabCap", "Table Caption")]);
    expect(m.figureCaption).toBe("FigCap");
    expect(m.tableCaption).toBe("TabCap");
  });
  it("body uses Normal first, then Body Text", () => {
    expect(buildStyleMap([p("BodyText", "Body Text"), p("Normal", "Normal")]).body).toBe("Normal");
    expect(buildStyleMap([p("BodyText", "Body Text")]).body).toBe("BodyText");
  });
  it("code accepts Code and HTML Preformatted", () => {
    expect(buildStyleMap([p("Code", "Code")]).code).toBe("Code");
    expect(buildStyleMap([p("HTMLPreformatted", "HTML Preformatted")]).code).toBe("HTMLPreformatted");
  });
  it("ignores character, table and numbering styles with the same name", () => {
    const m = buildStyleMap([
      { id: "Heading1", name: "heading 1", kind: "character" },
      { id: "Caption", name: "caption", kind: "table" },
    ]);
    expect(m.headings[0]).toBeNull();
    expect(m.figureCaption).toBeNull();
  });
  it("empty template gives all nulls", () => {
    const m = buildStyleMap([]);
    expect(m.title).toBeNull();
    expect(m.body).toBeNull();
    expect(m.headings.every((h) => h === null)).toBe(true);
  });
});

describe("headingStyleFor", () => {
  const m = buildStyleMap(FULL);
  it("exact level", () => {
    expect(headingStyleFor(m, 2)).toBe("Heading2");
  });
  it("deeper than the template uses the deepest one", () => {
    expect(headingStyleFor(m, 6)).toBe("Heading3");
  });
  it("clamps levels below 1 and above 9", () => {
    expect(headingStyleFor(m, 0)).toBe("Heading1");
    expect(headingStyleFor(m, 99)).toBe("Heading3");
  });
  it("null when the template has no heading at or above the level", () => {
    expect(headingStyleFor(buildStyleMap([p("Heading2", "heading 2")]), 1)).toBeNull();
    expect(headingStyleFor(buildStyleMap([]), 1)).toBeNull();
  });
});

describe("effectiveRole and styleIdForParagraph", () => {
  const m = buildStyleMap(FULL);
  it("manual role wins", () => {
    expect(effectiveRole(detected("body", null, "title"))).toEqual({ role: "title", level: null });
  });
  it("manual heading without a level uses the detected level, else 1", () => {
    expect(effectiveRole(detected("unknown", 3, "heading"))).toEqual({ role: "heading", level: 3 });
    expect(effectiveRole(detected("unknown", null, "heading"))).toEqual({ role: "heading", level: 1 });
  });
  it("heading level is clamped", () => {
    expect(effectiveRole(detected("heading", 20)).level).toBe(9);
    expect(effectiveRole(detected("heading", 0)).level).toBe(1);
  });
  it("maps each role", () => {
    const cases: [StructureRole, number | null, string | null][] = [
      ["title", null, "Title"],
      ["heading", 2, "Heading2"],
      ["body", null, "Normal"],
      ["list-item", null, "ListParagraph"],
      ["figure-caption", null, "Caption"],
      ["table-caption", null, "Caption"],
      ["quote", null, "Quote"],
      ["code", null, "SourceCode"],
      ["table", null, null],
      ["unknown", null, null],
    ];
    for (const [role, level, expected] of cases) {
      expect(styleIdForRole(m, role, level)).toBe(expected);
      expect(styleIdForParagraph(m, detected(role, level))).toBe(expected);
    }
  });
  it("manual tag changes the style", () => {
    expect(styleIdForParagraph(m, detected("body", null, "quote"))).toBe("Quote");
  });
  it("role without a template style gives null", () => {
    expect(styleIdForParagraph(buildStyleMap([p("Normal", "Normal")]), detected("title", null))).toBeNull();
  });
});

describe("missingStyleNames", () => {
  it("lists each missing name once, sorted", () => {
    const m = buildStyleMap([p("Normal", "Normal")]);
    const list = missingStyleNames(m, [
      detected("heading", 2),
      detected("heading", 2),
      detected("figure-caption", null),
      detected("table-caption", null),
      detected("title", null),
      detected("body", null),
      detected("table", null),
      detected("unknown", null),
    ]);
    expect(list).toEqual(["caption", "heading 2", "title"]);
  });
  it("empty when the template has everything", () => {
    const m = buildStyleMap(FULL);
    expect(missingStyleNames(m, [detected("heading", 5), detected("code", null), detected("list-item", null)])).toEqual([]);
  });
  it("reports list, code and body names", () => {
    const m = buildStyleMap([]);
    expect(missingStyleNames(m, [detected("list-item", null), detected("code", null), detected("body", null), detected("quote", null)])).toEqual([
      "list paragraph",
      "normal",
      "quote",
      "source code",
    ]);
  });
});
