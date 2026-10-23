import { describe, expect, it } from "vitest";
import {
  chainTransforms,
  createNumIdRemapper,
  listNumIds,
  mergeNumbering,
  parseNumbering,
  remapNumIds,
} from "./numbering_merge";
import { RawTransform, rewriteXmlString } from "./xml_rewrite";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function abs(id: number, extra = ""): string {
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>${extra}<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>`;
}
function num(id: number, abstractId: number, extra = ""): string {
  return `<w:num w:numId="${id}"><w:abstractNumId w:val="${abstractId}"/>${extra}</w:num>`;
}
function doc(body: string, rootExtra = ""): string {
  return `${DECL}<w:numbering ${W}${rootExtra}>\n  ${body}\n</w:numbering>`;
}

const TEMPLATE = doc(abs(0) + "\n  " + abs(1) + "\n  " + num(1, 0) + "\n  " + num(2, 1));
const SOURCE = doc(abs(0, '<w:pStyle w:val="Heading1"/>') + abs(5) + num(1, 0) + num(2, 5, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>'));

describe("parseNumbering", () => {
  it("splits root and children and reads ids", () => {
    const p = parseNumbering(TEMPLATE);
    expect(p.rootOpen?.name).toBe("w:numbering");
    expect(p.children.map((c) => c.kind)).toEqual(["abstractNum", "abstractNum", "num", "num"]);
    expect(p.children.map((c) => c.id)).toEqual([0, 1, 1, 2]);
    expect(p.children[3]?.abstractNumId).toBe(1);
  });
  it("classifies picture bullets and other children", () => {
    const x = doc('<w:numPicBullet w:numPicBulletId="0"><w:pict/></w:numPicBullet>' + abs(0) + num(1, 0) + '<w:numIdMacAtCleanup w:val="2"/>');
    const p = parseNumbering(x);
    expect(p.children.map((c) => c.kind)).toEqual(["numPicBullet", "abstractNum", "num", "other"]);
    expect(p.children[3]?.id).toBeUndefined();
  });
  it("keeps child XML intact including escaped attributes", () => {
    const x = doc('<w:abstractNum w:abstractNumId="0"><w:name w:val="A &amp; B"/></w:abstractNum>');
    expect(parseNumbering(x).children[0]?.xml).toContain('w:val="A &amp; B"');
  });
  it("handles an empty root", () => {
    expect(parseNumbering(`${DECL}<w:numbering ${W}/>`).children).toEqual([]);
  });
  it("throws on damaged XML and on DOCTYPE", () => {
    expect(() => parseNumbering("<w:numbering><w:num></w:numbering>")).toThrow();
    expect(() => parseNumbering('<!DOCTYPE x [<!ENTITY a "b">]><w:numbering/>')).toThrow(/DOCTYPE/);
  });
  it("lists numIds", () => {
    expect([...listNumIds(TEMPLATE)]).toEqual([1, 2]);
  });
});

describe("mergeNumbering", () => {
  it("returns undefined when neither side has numbering", () => {
    expect(mergeNumbering(undefined, undefined).xml).toBeUndefined();
  });
  it("returns the template unchanged when the source has none", () => {
    const m = mergeNumbering(TEMPLATE, undefined);
    expect(m.xml).toBe(TEMPLATE);
    expect(m.numIdMap.size).toBe(0);
  });
  it("shifts source ids and updates abstractNumId", () => {
    const m = mergeNumbering(TEMPLATE, SOURCE);
    expect([...m.numIdMap]).toEqual([[1, 3], [2, 4]]);
    expect([...m.abstractIdMap]).toEqual([[0, 2], [5, 7]]);
    const p = parseNumbering(m.xml ?? "");
    expect(p.children.map((c) => `${c.kind}:${c.id}`)).toEqual([
      "abstractNum:0", "abstractNum:1", "abstractNum:2", "abstractNum:7",
      "num:1", "num:2", "num:3", "num:4",
    ]);
    expect(p.children.filter((c) => c.kind === "num").map((c) => c.abstractNumId)).toEqual([0, 1, 2, 7]);
  });
  it("keeps template children byte for byte and source lvlOverride", () => {
    const m = mergeNumbering(TEMPLATE, SOURCE);
    for (const c of parseNumbering(TEMPLATE).children) expect(m.xml).toContain(c.xml);
    expect(m.xml).toContain('<w:startOverride w:val="3"/>');
  });
  it("removes source list-level style links and warns, template links stay", () => {
    const tpl = doc(abs(0, '<w:pStyle w:val="TplHeading1"/>') + num(1, 0));
    const m = mergeNumbering(tpl, SOURCE);
    expect(m.xml).not.toContain('w:val="Heading1"');
    expect(m.xml).toContain('<w:pStyle w:val="TplHeading1"/>');
    expect(m.warnings.join(" ")).toMatch(/Removed 1 style link/);
  });
  it("shifts picture bullets and their level references", () => {
    const pic = (id: number) => `<w:numPicBullet w:numPicBulletId="${id}"><w:pict/></w:numPicBullet>`;
    const tpl = doc(pic(0) + pic(1) + abs(0) + num(1, 0));
    const src = doc(pic(0) + abs(0, '<w:lvlPicBulletId w:val="0"/>') + num(1, 0));
    const m = mergeNumbering(tpl, src);
    expect([...m.picBulletIdMap]).toEqual([[0, 2]]);
    expect(m.xml).toContain('w:numPicBulletId="2"');
    expect(m.xml).toContain('<w:lvlPicBulletId w:val="2"/>');
    const p = parseNumbering(m.xml ?? "");
    expect(p.children.map((c) => c.kind)).toEqual(["numPicBullet", "numPicBullet", "numPicBullet", "abstractNum", "abstractNum", "num", "num"]);
  });
  it("uses shift 0 when the template has no numbering part", () => {
    const m = mergeNumbering(undefined, SOURCE);
    expect([...m.numIdMap]).toEqual([[1, 1], [2, 2]]);
    expect(listNumIds(m.xml ?? "")).toEqual(new Set([1, 2]));
  });
  it("uses larger shifts when template ids are high", () => {
    const tpl = doc(abs(4) + num(12, 4));
    const m = mergeNumbering(tpl, doc(abs(0) + num(1, 0)));
    expect(m.numIdMap.get(1)).toBe(13);
    expect(m.abstractIdMap.get(0)).toBe(5);
  });
  it("keeps template numIdMacAtCleanup and drops the source one", () => {
    const tpl = doc(abs(0) + num(1, 0) + '<w:numIdMacAtCleanup w:val="1"/>');
    const src = doc(abs(0) + num(1, 0) + '<w:numIdMacAtCleanup w:val="9"/>');
    const m = mergeNumbering(tpl, src);
    expect(m.xml).toContain('<w:numIdMacAtCleanup w:val="1"/>');
    expect(m.xml).not.toContain('w:val="9"');
    expect((m.xml ?? "").trim().endsWith("</w:numbering>")).toBe(true);
  });
  it("unites namespaces and mc:Ignorable", () => {
    const tpl = doc(abs(0) + num(1, 0), ' xmlns:mc="urn:mc" mc:Ignorable="w14"');
    const src = doc(abs(0) + num(1, 0), ' xmlns:mc="urn:mc" xmlns:w15="urn:w15" mc:Ignorable="w14 w15"');
    const m = mergeNumbering(tpl, src);
    expect(m.xml).toContain('xmlns:w15="urn:w15"');
    expect(m.xml).toContain('mc:Ignorable="w14 w15"');
  });
  it("warns when a namespace prefix means something else", () => {
    const tpl = doc(abs(0) + num(1, 0), ' xmlns:x="urn:a"');
    const src = doc(abs(0) + num(1, 0), ' xmlns:x="urn:b"');
    expect(mergeNumbering(tpl, src).warnings.join(" ")).toMatch(/xmlns:x/);
  });
  it("warns about style links in source list definitions", () => {
    const src = doc('<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="MyList"/></w:abstractNum>' + num(1, 0));
    expect(mergeNumbering(TEMPLATE, src).warnings.join(" ")).toMatch(/style link/);
  });
  it("is deterministic and does not touch its inputs", () => {
    const a = mergeNumbering(TEMPLATE, SOURCE);
    const b = mergeNumbering(TEMPLATE, SOURCE);
    expect(a.xml).toBe(b.xml);
    expect(TEMPLATE).toContain('w:abstractNumId="0"');
  });
  it("produces ids that are unique and re-parse cleanly", () => {
    const m = mergeNumbering(TEMPLATE, SOURCE);
    const kids = parseNumbering(m.xml ?? "").children;
    const nums = kids.filter((c) => c.kind === "num").map((c) => c.id);
    expect(new Set(nums).size).toBe(nums.length);
    const abstracts = kids.filter((c) => c.kind === "abstractNum").map((c) => c.id);
    expect(new Set(abstracts).size).toBe(abstracts.length);
  });
  it("merging an empty source changes nothing important", () => {
    const m = mergeNumbering(TEMPLATE, `${DECL}<w:numbering ${W}/>`);
    expect(parseNumbering(m.xml ?? "").children.length).toBe(4);
    expect(m.numIdMap.size).toBe(0);
  });
  it("throws for damaged source numbering", () => {
    expect(() => mergeNumbering(TEMPLATE, "<w:numbering><w:num></w:numbering>")).toThrow();
  });
});

describe("numId remapping", () => {
  const map = new Map<number, number>([[1, 3], [2, 4]]);
  const P = (id: number) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr></w:pPr></w:p>`;
  it("rewrites mapped ids, keeps 0 and counts unmapped", () => {
    const r = remapNumIds(`<w:body ${W}>${P(1)}${P(2)}${P(0)}${P(9)}</w:body>`, "word/document.xml", map);
    expect(r.xml).toContain('<w:numId w:val="3"/>');
    expect(r.xml).toContain('<w:numId w:val="4"/>');
    expect(r.xml).toContain('<w:numId w:val="0"/>');
    expect(r.xml).toContain('<w:numId w:val="9"/>');
    expect(r.stats).toEqual({ remapped: 2, unmapped: 1 });
  });
  it("also rewrites numPr inside tracked paragraph changes", () => {
    const x = `<w:p ${W}><w:pPr><w:pPrChange w:id="1"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:pPrChange></w:pPr></w:p>`;
    expect(remapNumIds(x, "p", map).xml).toContain('<w:numId w:val="3"/>');
  });
  it("does not touch other elements or the w:num numId attribute", () => {
    const x = `<w:r ${W}><w:t>1</w:t><w:num w:numId="1"/></w:r>`;
    expect(remapNumIds(x, "p", map).xml).toBe(x);
  });
  it("leaves everything alone with an empty map", () => {
    const x = `<w:body ${W}>${P(1)}</w:body>`;
    const r = remapNumIds(x, "p", new Map());
    expect(r.xml).toBe(x);
    expect(r.stats).toEqual({ remapped: 0, unmapped: 1 });
  });
  it("chains with another transform", () => {
    const upper: RawTransform = (t, emit) => emit(t.kind === "text" ? { kind: "text", text: t.text.toUpperCase() } : t);
    const remapper = createNumIdRemapper(map);
    const out = rewriteXmlString(`<w:p ${W}><w:t>abc</w:t><w:numId w:val="1"/></w:p>`, "p", chainTransforms(remapper.transform, upper));
    expect(out).toContain("ABC");
    expect(out).toContain('w:val="3"');
  });
  it("works with a different prefix", () => {
    expect(remapNumIds('<x:numId xmlns:x="urn:x" x:val="2"/>', "p", map).xml).toContain('x:val="4"');
  });
  it("is stable when run twice with a disjoint map (no double shift for 1->3 only once)", () => {
    const once = remapNumIds(`<w:body ${W}>${P(1)}</w:body>`, "p", map).xml;
    expect(once).toContain('w:val="3"');
    expect(once).not.toContain('w:val="1"');
  });
});
