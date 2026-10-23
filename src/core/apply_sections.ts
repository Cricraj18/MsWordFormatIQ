/**
 * formatIQ apply helpers for section properties and settings (part of file 3b-4).
 *
 * - readSectionModel: reads, from the TEMPLATE document.xml, the header/footer
 *   references and the titlePg flag of the first section that has references.
 * - createSectionFixer: a token transform for the OUTPUT document.xml. Every
 *   w:sectPr loses the SOURCE header/footer references (they would dangle) and
 *   gets the template's references (with the new relationship ids).
 *   w:titlePg is taken from the template model only.
 * - addUpdateFields: puts w:updateFields into settings.xml at the right place.
 *
 * Limit (documented): all source sections get the same references (the
 * template's first section with references). Per-section choices come in 5a/5b.
 */
import { closeTag, openTag, rewriteXmlString, withAttr } from "./xml_rewrite";
import type { RawAttr, RawOpen, RawToken, RawTransform } from "./xml_rewrite";

/** Namespace of r:id attributes. */
export const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** One header or footer reference of a section. */
export interface SectionRef {
  kind: "header" | "footer";
  /** w:type value: "default", "first" or "even". */
  type: string;
  /** Relationship id (in the TEMPLATE document.xml.rels). */
  id: string;
}

/** What the template's section says about headers and footers. */
export interface SectionModel {
  refs: SectionRef[];
  titlePg: boolean;
}

/** Model with nothing in it. */
export function emptySectionModel(): SectionModel {
  return { refs: [], titlePg: false };
}

function localOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function prefixOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? "" : name.slice(0, colon + 1);
}

function attrLocal(tag: RawOpen, local: string): string | undefined {
  for (const [key, value] of tag.attrs) {
    if (!key.startsWith("xmlns") && localOf(key) === local) return value;
  }
  return undefined;
}

function isOn(value: string | undefined): boolean {
  if (value === undefined) return true;
  const v = value.toLowerCase();
  return !(v === "0" || v === "false" || v === "off");
}

/** Reader for the template's section model. Feed document.xml tokens through transform. */
export interface SectionModelReader {
  transform: RawTransform;
  getModel(): SectionModel;
}

/** Make a reader. The first sectPr that has header/footer references is the model. */
export function createSectionModelReader(): SectionModelReader {
  const stack: string[] = [];
  let changeDepth = 0;
  let done = false;
  let collecting = false;
  let sectLen = -1;
  let refs: SectionRef[] = [];
  let titlePg = false;
  let model: SectionModel = emptySectionModel();

  const transform: RawTransform = (token) => {
    if (token.kind === "open") {
      const local = localOf(token.name);
      const parent = stack.length > 0 ? localOf(stack[stack.length - 1] ?? "") : "";
      if (local === "sectPrChange") changeDepth += 1;
      if (!done && changeDepth === 0 && local === "sectPr" && !collecting) {
        collecting = true;
        refs = [];
        titlePg = false;
        stack.push(token.name);
        sectLen = stack.length;
        return;
      }
      if (collecting && parent === "sectPr" && stack.length === sectLen) {
        if (local === "headerReference" || local === "footerReference") {
          const id = attrLocal(token, "id");
          if (id !== undefined) {
            refs.push({ kind: local === "headerReference" ? "header" : "footer", type: attrLocal(token, "type") ?? "default", id });
          }
        } else if (local === "titlePg") {
          titlePg = isOn(attrLocal(token, "val"));
        }
      }
      stack.push(token.name);
      return;
    }
    if (token.kind === "close") {
      const local = localOf(token.name);
      if (collecting && local === "sectPr" && stack.length === sectLen) {
        collecting = false;
        sectLen = -1;
        if (refs.length > 0) {
          model = { refs, titlePg };
          done = true;
        }
      }
      if (local === "sectPrChange") changeDepth -= 1;
      stack.pop();
    }
  };
  return { transform, getModel: () => model };
}

/** Counters of the section fixer. */
export interface SectionFixStats {
  /** Section property blocks changed. */
  sections: number;
  /** Source header/footer references removed (also inside tracked section changes). */
  removedSourceRefs: number;
  /** Template references written. */
  insertedRefs: number;
}

/** Elements that come AFTER w:titlePg in the schema order of w:sectPr. */
const AFTER_TITLE_PG: ReadonlySet<string> = new Set([
  "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]);

interface SectFrame {
  prefix: string;
  active: boolean;
  titleDone: boolean;
}

/**
 * Transform for the output document.xml. templateIdMap maps the template's
 * relationship ids to the ids used in the output (from mergeRelationships).
 * Adds xmlns:r to the root element when references are written and the root
 * has no such declaration.
 */
export function createSectionFixer(
  model: SectionModel,
  templateIdMap: ReadonlyMap<string, string>,
): { transform: RawTransform; stats: SectionFixStats } {
  const stats: SectionFixStats = { sections: 0, removedSourceRefs: 0, insertedRefs: 0 };
  const refs = model.refs.flatMap((ref) => {
    const id = templateIdMap.get(ref.id);
    return id === undefined ? [] : [{ ...ref, id }];
  });
  const stack: string[] = [];
  const frames: SectFrame[] = [];
  let skipDepth = 0;
  let changeDepth = 0;
  let rootSeen = false;

  const writeTitlePg = (frame: SectFrame, emit: (token: RawToken) => void): void => {
    if (frame.active && model.titlePg && !frame.titleDone) {
      emit(openTag(frame.prefix + "titlePg"));
      emit(closeTag(frame.prefix + "titlePg"));
      frame.titleDone = true;
    }
  };

  const transform: RawTransform = (token, emit) => {
    if (skipDepth > 0) {
      if (token.kind === "open") {
        stack.push(token.name);
        skipDepth += 1;
      } else if (token.kind === "close") {
        stack.pop();
        skipDepth -= 1;
      }
      return;
    }
    if (token.kind === "open") {
      let tag: RawOpen = token;
      if (!rootSeen) {
        rootSeen = true;
        if (refs.length > 0 && !tag.attrs.some((a: RawAttr) => a[0] === "xmlns:r")) {
          tag = withAttr(tag, "xmlns:r", OFFICE_REL_NS);
        }
      }
      const local = localOf(tag.name);
      const parent = stack.length > 0 ? localOf(stack[stack.length - 1] ?? "") : "";
      const frame = frames[frames.length - 1];

      if (local === "sectPr") {
        const active = changeDepth === 0;
        emit(tag);
        stack.push(tag.name);
        const prefix = prefixOf(tag.name);
        frames.push({ prefix, active, titleDone: false });
        if (active) {
          stats.sections += 1;
          for (const ref of refs) {
            emit(openTag(prefix + (ref.kind === "header" ? "headerReference" : "footerReference"), [[prefix + "type", ref.type], ["r:id", ref.id]]));
            emit(closeTag(prefix + (ref.kind === "header" ? "headerReference" : "footerReference")));
            stats.insertedRefs += 1;
          }
        }
        return;
      }
      if (frame !== undefined && parent === "sectPr") {
        if (local === "headerReference" || local === "footerReference") {
          stats.removedSourceRefs += 1;
          stack.push(tag.name);
          skipDepth = 1;
          return;
        }
        if (local === "titlePg") {
          stack.push(tag.name);
          skipDepth = 1;
          return;
        }
        if (AFTER_TITLE_PG.has(local)) writeTitlePg(frame, emit);
      }
      if (local === "sectPrChange") changeDepth += 1;
      emit(tag);
      stack.push(tag.name);
      return;
    }
    if (token.kind === "close") {
      const local = localOf(token.name);
      if (local === "sectPr") {
        const frame = frames.pop();
        if (frame !== undefined) writeTitlePg(frame, emit);
      }
      if (local === "sectPrChange") changeDepth -= 1;
      stack.pop();
      emit(token);
      return;
    }
    emit(token);
  };
  return { transform, stats };
}

/** Elements that come AFTER w:updateFields in the schema order of w:settings. */
const AFTER_UPDATE_FIELDS: ReadonlySet<string> = new Set([
  "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr",
  "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
  "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown",
  "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol",
  "listSeparator",
]);

/**
 * Make settings.xml ask Word to update fields when the file opens.
 * Adds w:updateFields w:val="true" in schema order, or turns an existing
 * w:updateFields on. Returns the new text and whether anything changed.
 */
export function addUpdateFields(settingsXml: string, partName = "word/settings.xml"): { xml: string; changed: boolean } {
  let exists = false;
  let alreadyOn = false;
  let depth = 0;
  rewriteXmlString(settingsXml, partName, (token) => {
    if (token.kind === "open") {
      depth += 1;
      if (depth === 2 && localOf(token.name) === "updateFields") {
        exists = true;
        alreadyOn = isOn(attrLocal(token, "val"));
      }
    } else if (token.kind === "close") {
      depth -= 1;
    }
  });
  if (exists && alreadyOn) return { xml: settingsXml, changed: false };

  depth = 0;
  let inserted = false;
  const xml = rewriteXmlString(settingsXml, partName, (token, emit) => {
    if (token.kind === "open") {
      depth += 1;
      const local = localOf(token.name);
      if (depth === 2 && exists && local === "updateFields") {
        emit(withAttr(token, prefixOf(token.name) + "val", "true"));
        return;
      }
      if (depth === 2 && !exists && !inserted && AFTER_UPDATE_FIELDS.has(local)) {
        const p = prefixOf(token.name);
        emit(openTag(p + "updateFields", [[p + "val", "true"]]));
        emit(closeTag(p + "updateFields"));
        inserted = true;
      }
      emit(token);
      return;
    }
    if (token.kind === "close") {
      if (depth === 1 && !exists && !inserted) {
        const p = prefixOf(token.name);
        emit(openTag(p + "updateFields", [[p + "val", "true"]]));
        emit(closeTag(p + "updateFields"));
        inserted = true;
      }
      depth -= 1;
    }
    emit(token);
  });
  return { xml, changed: true };
}
