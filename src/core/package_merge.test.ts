import { describe, expect, it } from "vitest";
import { FormatIQError } from "../types";
import {
  collectPartClosure,
  contentTypeOf,
  isTemplateOwned,
  mergeContentTypes,
  mergeRelationships,
  parseContentTypes,
  parseRelationships,
  partDir,
  planPartNames,
  pruneOverrides,
  relativeTarget,
  relsPathOf,
  resolveTarget,
  retargetRels,
  writeContentTypes,
  writeRelationships,
} from "./package_merge";
import type { Relationship } from "./package_merge";

const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const rel = (id: string, kind: string, target: string, targetMode?: string): Relationship =>
  targetMode === undefined ? { id, type: T + kind, target } : { id, type: T + kind, target, targetMode };

describe("relationship files", () => {
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${T}styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="${T}hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>` +
    `<Relationship Id="rId3" Type="${T}image" Target="media/image1.png"/></Relationships>`;
  it("parses ids, types, targets and target mode (entities decoded)", () => {
    const list = parseRelationships(xml);
    expect(list).toHaveLength(3);
    expect(list[0]).toEqual({ id: "rId1", type: T + "styles", target: "styles.xml" });
    expect(list[1]).toEqual({ id: "rId2", type: T + "hyperlink", target: "https://example.com/?a=1&b=2", targetMode: "External" });
  });
  it("writes and re-reads the same list (escaping the ampersand)", () => {
    const list = parseRelationships(xml);
    const written = writeRelationships(list);
    expect(written).toContain("a=1&amp;b=2");
    expect(parseRelationships(written)).toEqual(list);
  });
  it("skips elements that lack Id, Type or Target", () => {
    expect(parseRelationships(`<Relationships><Relationship Id="x" Type="y"/><Relationship Type="a" Target="b"/></Relationships>`)).toEqual([]);
  });
  it("empty list writes a valid empty file", () => {
    expect(parseRelationships(writeRelationships([]))).toEqual([]);
  });
  it("damaged XML throws ZIP_READ_FAILED", () => {
    expect(() => parseRelationships("<Relationships><Relationship>")).toThrow(FormatIQError);
  });
});

describe("template-owned types", () => {
  it("styles, theme, fontTable, settings, webSettings, numbering, header, footer", () => {
    for (const k of ["styles", "theme", "fontTable", "settings", "webSettings", "numbering", "header", "footer"]) {
      expect(isTemplateOwned(T + k)).toBe(true);
    }
  });
  it("images, hyperlinks, footnotes, comments are not", () => {
    for (const k of ["image", "hyperlink", "footnotes", "endnotes", "comments", "chart", "oleObject", "customXml"]) {
      expect(isTemplateOwned(T + k)).toBe(false);
    }
  });
});

describe("mergeRelationships", () => {
  const source = [
    rel("rId1", "styles", "styles.xml"),
    rel("rId2", "numbering", "numbering.xml"),
    rel("rId3", "image", "media/image1.png"),
    rel("rId4", "hyperlink", "https://x.test", "External"),
    rel("rId5", "footnotes", "footnotes.xml"),
    rel("rId6", "header", "header1.xml"),
  ];
  const template = [
    rel("rId1", "styles", "styles.xml"),
    rel("rId2", "theme", "theme/theme1.xml"),
    rel("rId3", "header", "header1.xml"),
    rel("rId4", "footer", "footer1.xml"),
    rel("rId5", "image", "media/logo.png"),
  ];
  const merged = mergeRelationships(source, template);
  it("keeps non-owned source relationships with their own ids", () => {
    expect(merged.rels.filter((r) => r.id.startsWith("rId")).map((r) => r.id)).toEqual(["rId3", "rId4", "rId5"]);
  });
  it("drops source relationships of owned types and lists them", () => {
    expect(merged.droppedSourceRels.map((r) => r.id)).toEqual(["rId1", "rId2", "rId6"]);
  });
  it("takes owned template relationships with new ids and a map", () => {
    expect([...merged.templateIdMap.entries()]).toEqual([
      ["rId1", "tplrId1"],
      ["rId2", "tplrId2"],
      ["rId3", "tplrId3"],
      ["rId4", "tplrId4"],
    ]);
    expect(merged.rels.find((r) => r.id === "tplrId3")?.target).toBe("header1.xml");
  });
  it("does not take template images (not owned)", () => {
    expect(merged.rels.some((r) => r.target === "media/logo.png")).toBe(false);
  });
  it("all ids in the result are unique", () => {
    const ids = merged.rels.map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
  it("avoids a clash when the source already uses a tpl id", () => {
    const m = mergeRelationships([rel("tplrId1", "image", "media/a.png")], [rel("rId1", "styles", "styles.xml")]);
    const ids = m.rels.map((r) => r.id);
    expect(new Set(ids).size).toBe(2);
    expect(m.templateIdMap.get("rId1")).toBe("tpl2_rId1");
  });
  it("empty inputs give an empty result", () => {
    const m = mergeRelationships([], []);
    expect(m.rels).toEqual([]);
    expect(m.templateIdMap.size).toBe(0);
  });
});

describe("part names and targets", () => {
  it("partDir and relsPathOf", () => {
    expect(partDir("word/document.xml")).toBe("word");
    expect(partDir("[Content_Types].xml")).toBe("");
    expect(relsPathOf("word/footnotes.xml")).toBe("word/_rels/footnotes.xml.rels");
    expect(relsPathOf("root.xml")).toBe("_rels/root.xml.rels");
  });
  it("resolveTarget: relative, parent, absolute, dot segments", () => {
    expect(resolveTarget("word", "media/a.png")).toBe("word/media/a.png");
    expect(resolveTarget("word", "../customXml/item1.xml")).toBe("customXml/item1.xml");
    expect(resolveTarget("word", "/word/theme/theme1.xml")).toBe("word/theme/theme1.xml");
    expect(resolveTarget("word", "./a/./b.xml")).toBe("word/a/b.xml");
    expect(resolveTarget("", "word/document.xml")).toBe("word/document.xml");
  });
  it("resolveTarget returns undefined when .. climbs out of the zip", () => {
    expect(resolveTarget("word", "../../evil.xml")).toBeUndefined();
    expect(resolveTarget("", "../x")).toBeUndefined();
  });
  it("relativeTarget", () => {
    expect(relativeTarget("word", "word/media/a.png")).toBe("media/a.png");
    expect(relativeTarget("word", "customXml/item1.xml")).toBe("../customXml/item1.xml");
    expect(relativeTarget("word/embeddings", "word/media/a.png")).toBe("../media/a.png");
    expect(relativeTarget("", "word/document.xml")).toBe("word/document.xml");
    expect(relativeTarget("word", "word/footnotes.xml")).toBe("footnotes.xml");
  });
  it("resolve and relative are inverse", () => {
    for (const [dir, name] of [["word", "word/media/a.png"], ["word", "customXml/i.xml"], ["word/x/y", "word/z.xml"]] as const) {
      expect(resolveTarget(dir, relativeTarget(dir, name))).toBe(name);
    }
  });
});

describe("collectPartClosure", () => {
  const roots = [rel("r1", "image", "media/a.png"), rel("r2", "footnotes", "footnotes.xml"), rel("r3", "hyperlink", "https://x.test", "External")];
  it("follows relationships of reached parts and ignores external ones", () => {
    const relsByPart = new Map<string, Relationship[]>([
      ["word/footnotes.xml", [rel("f1", "image", "media/b.png"), rel("f2", "hyperlink", "http://y.test", "External")]],
    ]);
    expect(collectPartClosure(roots, "word", relsByPart)).toEqual(["word/media/a.png", "word/footnotes.xml", "word/media/b.png"]);
  });
  it("lists a shared part once and survives cycles", () => {
    const relsByPart = new Map<string, Relationship[]>([
      ["word/footnotes.xml", [rel("f1", "footnotes", "endnotes.xml"), rel("f2", "image", "media/a.png")]],
      ["word/endnotes.xml", [rel("e1", "footnotes", "footnotes.xml")]],
    ]);
    const closure = collectPartClosure([rel("r1", "footnotes", "footnotes.xml"), rel("r2", "image", "media/a.png")], "word", relsByPart);
    expect(closure).toEqual(["word/footnotes.xml", "word/media/a.png", "word/endnotes.xml"]);
  });
  it("ignores targets that leave the zip", () => {
    expect(collectPartClosure([rel("r1", "image", "../../x.png")], "word", new Map())).toEqual([]);
  });
});

describe("planPartNames and retargetRels", () => {
  it("keeps names that do not clash", () => {
    const plan = planPartNames(["word/media/a.png"], ["word/media/b.png"]);
    expect(plan.get("word/media/a.png")).toBe("word/media/a.png");
  });
  it("renames a clash with src_ and numbers further clashes", () => {
    const plan = planPartNames(
      ["word/media/image1.png", "word/media/src_image1.png"],
      ["word/media/image1.png", "word/media/SRC_image1.png"],
    );
    expect(plan.get("word/media/image1.png")).toBe("word/media/src2_image1.png");
    expect(plan.get("word/media/src_image1.png")).toBe("word/media/src_src_image1.png");
  });
  it("names are compared ignoring case", () => {
    expect(planPartNames(["word/media/Image1.PNG"], ["word/media/image1.png"]).get("word/media/Image1.PNG")).toBe("word/media/src_Image1.PNG");
  });
  it("two source parts with the same name in different cases do not collide with each other", () => {
    const plan = planPartNames(["word/a.xml", "word/A.XML"], []);
    expect(new Set([...plan.values()].map((v) => v.toLowerCase())).size).toBe(2);
  });
  it("retargetRels follows renamed parts and keeps external ones", () => {
    const renames = new Map([["word/media/image1.png", "word/media/src_image1.png"]]);
    const out = retargetRels(
      [rel("a", "image", "media/image1.png"), rel("b", "hyperlink", "https://x.test", "External"), rel("c", "image", "media/other.png")],
      "word/document.xml",
      "word/document.xml",
      renames,
    );
    expect(out.map((r) => r.target)).toEqual(["media/src_image1.png", "https://x.test", "media/other.png"]);
  });
  it("retargetRels rewrites targets relative to the part's new folder", () => {
    const out = retargetRels([rel("a", "image", "../media/p.png")], "word/embeddings/x.xml", "word/x.xml", new Map());
    expect(out[0]?.target).toBe("media/p.png");
  });
});

describe("content types", () => {
  const templateXml =
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/word/gone.xml" ContentType="application/x-gone"/></Types>`;
  const sourceXml =
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Default Extension="emf" ContentType="image/x-emf"/>` +
    `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>` +
    `<Override PartName="/word/media/odd.png" ContentType="image/special"/></Types>`;
  const template = parseContentTypes(templateXml);
  const source = parseContentTypes(sourceXml);
  it("parses defaults (lower case extension) and overrides", () => {
    expect(template.defaults.get("png")).toBe("image/png");
    expect(template.overrides.get("/word/styles.xml")).toContain("styles+xml");
    expect(parseContentTypes(`<Types><Default Extension="PNG" ContentType="image/png"/></Types>`).defaults.has("png")).toBe(true);
  });
  it("write and parse round trip", () => {
    const again = parseContentTypes(writeContentTypes(template));
    expect([...again.defaults]).toEqual([...template.defaults]);
    expect([...again.overrides]).toEqual([...template.overrides]);
  });
  it("contentTypeOf prefers Override over Default", () => {
    expect(contentTypeOf(source, "word/media/odd.png")).toBe("image/special");
    expect(contentTypeOf(source, "word/media/x.png")).toBe("image/png");
    expect(contentTypeOf(source, "word/none.bin")).toBeUndefined();
  });
  it("adds a Default for a new extension, an Override for a part with its own type", () => {
    const copied = new Map([
      ["word/media/photo.jpeg", "word/media/photo.jpeg"],
      ["word/footnotes.xml", "word/footnotes.xml"],
      ["word/media/odd.png", "word/media/odd.png"],
      ["word/media/draw.emf", "word/media/draw.emf"],
    ]);
    const merged = mergeContentTypes(template, source, copied);
    expect(merged.defaults.get("jpeg")).toBe("image/jpeg");
    expect(merged.defaults.get("emf")).toBe("image/x-emf");
    expect(merged.overrides.get("/word/footnotes.xml")).toContain("footnotes+xml");
    expect(merged.overrides.get("/word/media/odd.png")).toBe("image/special");
    expect(merged.defaults.get("png")).toBe("image/png");
  });
  it("uses the output name of a renamed part", () => {
    const merged = mergeContentTypes(template, source, new Map([["word/media/odd.png", "word/media/src_odd.png"]]));
    expect(merged.overrides.get("/word/media/src_odd.png")).toBe("image/special");
    expect(merged.overrides.has("/word/media/odd.png")).toBe(false);
  });
  it("does not change the template's own entries and does not mutate its inputs", () => {
    const before = writeContentTypes(template);
    const merged = mergeContentTypes(template, source, new Map([["word/media/photo.jpeg", "word/media/photo.jpeg"]]));
    expect(writeContentTypes(template)).toBe(before);
    expect(merged.overrides.get("/word/document.xml")).toBe(template.overrides.get("/word/document.xml"));
  });
  it("skips source parts that have no content type", () => {
    const merged = mergeContentTypes(template, source, new Map([["word/mystery.bin", "word/mystery.bin"]]));
    expect(merged.defaults.has("bin")).toBe(false);
    expect(merged.overrides.has("/word/mystery.bin")).toBe(false);
  });
  it("pruneOverrides removes overrides of parts that are not in the output", () => {
    const pruned = pruneOverrides(template, ["word/document.xml", "word/styles.xml"]);
    expect(pruned.overrides.has("/word/gone.xml")).toBe(false);
    expect(pruned.overrides.has("/word/styles.xml")).toBe(true);
    expect(template.overrides.has("/word/gone.xml")).toBe(true);
    expect(pruned.defaults.size).toBe(template.defaults.size);
  });
});
