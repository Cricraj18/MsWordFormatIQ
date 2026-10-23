import { describe, expect, it } from "vitest";
import { rewriteXmlString } from "./xml_rewrite";
import { addUpdateFields, createSectionFixer, createSectionModelReader } from "./apply_sections";
import type { SectionModel } from "./apply_sections";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function readModel(xml: string): SectionModel {
  const reader = createSectionModelReader();
  rewriteXmlString(xml, "t.xml", reader.transform);
  return reader.getModel();
}

function fix(xml: string, model: SectionModel, ids: [string, string][]): { xml: string; sections: number; removed: number; inserted: number } {
  const fixer = createSectionFixer(model, new Map(ids));
  const out = rewriteXmlString(xml, "d.xml", fixer.transform);
  return { xml: out, sections: fixer.stats.sections, removed: fixer.stats.removedSourceRefs, inserted: fixer.stats.insertedRefs };
}

describe("createSectionModelReader", () => {
  it("takes the first sectPr that has references", () => {
    const xml = `<w:document ${W}><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:pPr></w:p><w:sectPr><w:headerReference w:type="first" r:id="rA"/><w:footerReference w:type="default" r:id="rB"/><w:titlePg/></w:sectPr></w:body></w:document>`;
    expect(readModel(xml)).toEqual({ refs: [{ kind: "header", type: "first", id: "rA" }, { kind: "footer", type: "default", id: "rB" }], titlePg: true });
  });
  it("gives an empty model when there are no references", () => {
    expect(readModel(`<w:document ${W}><w:body><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:body></w:document>`)).toEqual({ refs: [], titlePg: false });
  });
  it("reads titlePg val 0 as off and ignores sectPrChange", () => {
    const xml = `<w:document ${W}><w:body><w:sectPr><w:headerReference w:type="default" r:id="rA"/><w:titlePg w:val="0"/><w:sectPrChange w:id="1"><w:sectPr><w:headerReference w:type="even" r:id="rZ"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`;
    expect(readModel(xml)).toEqual({ refs: [{ kind: "header", type: "default", id: "rA" }], titlePg: false });
  });
});

describe("createSectionFixer", () => {
  const model: SectionModel = { refs: [{ kind: "header", type: "default", id: "t1" }], titlePg: true };
  it("replaces references, adds titlePg in schema order and xmlns:r when missing", () => {
    const xml = `<w:document ${W}><w:body><w:sectPr><w:headerReference w:type="default" r:id="old"/><w:pgSz w:w="1" w:h="2"/><w:docGrid w:linePitch="1"/></w:sectPr></w:body></w:document>`;
    const out = fix(xml, model, [["t1", "tplt1"]]);
    expect(out.xml).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    expect(out.xml).toContain('<w:sectPr><w:headerReference w:type="default" r:id="tplt1"/><w:pgSz w:w="1" w:h="2"/><w:titlePg/><w:docGrid w:linePitch="1"/></w:sectPr>');
    expect(out).toMatchObject({ sections: 1, removed: 1, inserted: 1 });
  });
  it("drops source titlePg when the template has none and handles a section without references", () => {
    const xml = `<w:document ${W}><w:body><w:sectPr><w:titlePg/><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:body></w:document>`;
    const out = fix(xml, { refs: [], titlePg: false }, []);
    expect(out.xml).toContain("<w:sectPr><w:pgSz w:w=\"1\" w:h=\"2\"/></w:sectPr>");
    expect(out.xml).not.toContain("xmlns:r");
  });
  it("skips template references whose id has no mapping", () => {
    const out = fix(`<w:document ${W}><w:body><w:sectPr/></w:body></w:document>`, model, []);
    expect(out.inserted).toBe(0);
  });
  it("removes references inside sectPrChange without adding new ones", () => {
    const xml = `<w:document ${W}><w:body><w:sectPr><w:sectPrChange w:id="1"><w:sectPr><w:headerReference w:type="default" r:id="old"/></w:sectPr></w:sectPrChange></w:sectPr></w:body></w:document>`;
    const out = fix(xml, { refs: [{ kind: "header", type: "default", id: "t1" }], titlePg: false }, [["t1", "x"]]);
    expect(out.xml).not.toContain('r:id="old"');
    expect(out.xml.match(/headerReference/g)?.length).toBe(1);
    expect(out.sections).toBe(1);
  });
  it("puts titlePg at the end when nothing follows it", () => {
    const out = fix(`<w:document ${W}><w:body><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:body></w:document>`, model, [["t1", "x"]]);
    expect(out.xml).toContain('<w:pgSz w:w="1" w:h="2"/><w:titlePg/></w:sectPr>');
  });
});

describe("addUpdateFields", () => {
  it("inserts before compat", () => {
    const r = addUpdateFields(`<w:settings ${W}><w:zoom w:percent="100"/><w:compat/></w:settings>`);
    expect(r.changed).toBe(true);
    expect(r.xml).toContain('<w:zoom w:percent="100"/><w:updateFields w:val="true"/><w:compat/>');
  });
  it("appends at the end when no later element exists", () => {
    expect(addUpdateFields(`<w:settings ${W}><w:zoom w:percent="100"/></w:settings>`).xml).toContain('<w:updateFields w:val="true"/></w:settings>');
  });
  it("leaves an existing on switch alone and turns an off switch on", () => {
    const on = `<w:settings ${W}><w:updateFields w:val="true"/></w:settings>`;
    expect(addUpdateFields(on)).toEqual({ xml: on, changed: false });
    const off = addUpdateFields(`<w:settings ${W}><w:compat/><w:updateFields w:val="0"/></w:settings>`);
    expect(off.changed).toBe(true);
    expect(off.xml).toBe(`<w:settings ${W}><w:compat/><w:updateFields w:val="true"/></w:settings>`);
  });
});
