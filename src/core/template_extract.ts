/**
 * formatIQ template reader (file 1a).
 *
 * Reads a template .docx and fills a TemplateInfo (see src/types.ts):
 *  - styles           from word/styles.xml
 *  - page setup       from the FIRST w:sectPr in word/document.xml
 *  - hasHeaderFooter  true if any word/header*.xml or word/footer*.xml exists
 *  - hasToc           true if a TOC field (or a "Table of Contents" content
 *                     control) is found in word/document.xml
 *
 * The template file is only read, never changed. Nothing is uploaded.
 * XML is streamed (see xml.ts), so a large template is fine.
 */
import { FormatIQError } from "../types";
import type { PageSetup, ProgressCallback, StyleKind, TemplateInfo, TemplateStyle } from "../types";
import { DEFAULT_ZIP_LIMITS, openDocx } from "./zipio";
import type { JobControl, ZipLimits } from "./zipio";
import { parseEntryXml } from "./xml";
import type { XmlHandlers } from "./xml";

/** What the caller must or may provide. */
export interface TemplateExtractOptions {
  /** Id used when the template is stored (the caller makes it unique). */
  id: string;
  /** Name shown to the user. */
  name: string;
  /** ISO date text. Defaults to now. */
  createdAt?: string;
  /** Safety limits. Defaults to DEFAULT_ZIP_LIMITS. */
  limits?: ZipLimits;
  /** Cancel button and time limit. */
  control?: JobControl;
  /** Progress updates for the progress bar. */
  onProgress?: ProgressCallback;
}

/** Used when the template has no page size: A4 portrait (twips). */
export const DEFAULT_PAGE: Readonly<PageSetup> = {
  widthTwips: 11906,
  heightTwips: 16838,
  landscape: false,
};

const STYLE_KINDS: ReadonlySet<string> = new Set(["paragraph", "character", "table", "numbering"]);

function isStyleKind(value: string | undefined): value is StyleKind {
  return value !== undefined && STYLE_KINDS.has(value);
}

function positiveInt(value: string | undefined): number | null {
  if (value === undefined || !/^\d+$/.test(value)) return null;
  const n = Number(value);
  return Number.isSafeInteger(n) && n > 0 ? n : null;
}

/** Read the styles from styles.xml text parts. Keeps the first of duplicate ids. */
function styleHandlers(found: Map<string, TemplateStyle>): XmlHandlers {
  let kind: StyleKind | null = null;
  let id = "";
  let name = "";
  return {
    open(tag) {
      if (tag.name === "style") {
        const type = tag.attrs["type"];
        kind = isStyleKind(type) ? type : null;
        id = tag.attrs["styleId"] ?? "";
        name = "";
      } else if (tag.name === "name" && kind !== null && name === "") {
        name = tag.attrs["val"] ?? "";
      }
    },
    close(tagName) {
      if (tagName === "style" && kind !== null && id !== "" && !found.has(id)) {
        found.set(id, { id, name: name === "" ? id : name, kind });
      }
      if (tagName === "style") kind = null;
    },
  };
}

/** Result collected while streaming document.xml. */
interface DocumentFacts {
  page: PageSetup | null;
  hasToc: boolean;
}

/**
 * Read the first w:sectPr page size and look for a TOC field.
 * Only a pgSz that is a DIRECT child of the first sectPr counts (a pgSz inside
 * w:sectPrChange is an old tracked value and is ignored).
 */
function documentHandlers(facts: DocumentFacts): XmlHandlers {
  let depth = 0;
  let sectDepth = -1; // depth of the first sectPr while it is open
  let sectDone = false;
  let pendingPage: PageSetup | null = null;
  let instr: string | null = null;

  return {
    open(tag) {
      depth += 1;
      switch (tag.name) {
        case "sectPr":
          if (!sectDone && sectDepth === -1) sectDepth = depth;
          break;
        case "pgSz":
          if (sectDepth !== -1 && depth === sectDepth + 1 && pendingPage === null) {
            const w = positiveInt(tag.attrs["w"]);
            const h = positiveInt(tag.attrs["h"]);
            if (w !== null && h !== null) {
              const orient = tag.attrs["orient"];
              pendingPage = {
                widthTwips: w,
                heightTwips: h,
                landscape: orient === undefined ? w > h : orient === "landscape",
              };
            }
          }
          break;
        case "instrText":
          instr = "";
          break;
        case "fldSimple":
          if (/^\s*TOC\b/i.test(tag.attrs["instr"] ?? "")) facts.hasToc = true;
          break;
        case "docPartGallery":
          if ((tag.attrs["val"] ?? "").toLowerCase() === "table of contents") facts.hasToc = true;
          break;
        default:
          break;
      }
    },
    text(text) {
      if (instr !== null) instr += text;
    },
    close(tagName) {
      if (tagName === "instrText" && instr !== null) {
        if (/^\s*TOC\b/i.test(instr)) facts.hasToc = true;
        instr = null;
      }
      if (tagName === "sectPr" && sectDepth === depth) {
        facts.page = pendingPage;
        sectDone = true;
        sectDepth = -1;
      }
      depth -= 1;
    },
  };
}

/**
 * Read a template .docx.
 * @throws FormatIQError NOT_DOCX, TEMPLATE_INVALID (no styles.xml),
 *         XML_ENTITY_EXPANSION, ZIP_READ_FAILED, FILE_TOO_LARGE,
 *         CANCELLED, TIMEOUT.
 */
export async function extractTemplateInfo(
  source: Blob,
  fileName: string,
  options: TemplateExtractOptions,
): Promise<TemplateInfo> {
  const limits = options.limits ?? DEFAULT_ZIP_LIMITS;
  const report = (fraction: number, message: string): void => {
    options.onProgress?.({ phase: "reading", fraction, message });
  };

  report(0, "Opening template");
  const docx = await openDocx(source, fileName, limits, options.control);
  try {
    const stylesEntry = docx.entries.get("word/styles.xml");
    if (stylesEntry === undefined) {
      throw new FormatIQError("TEMPLATE_INVALID", "The template has no styles (word/styles.xml is missing).");
    }
    report(0.1, "Reading styles");
    const found = new Map<string, TemplateStyle>();
    await parseEntryXml(stylesEntry, styleHandlers(found), limits.maxXmlBytes, options.control);

    const mainEntry = docx.entries.get("word/document.xml");
    if (mainEntry === undefined) {
      throw new FormatIQError("NOT_DOCX", "This does not look like a Word document.");
    }
    report(0.4, "Reading page setup");
    const facts: DocumentFacts = { page: null, hasToc: false };
    await parseEntryXml(mainEntry, documentHandlers(facts), limits.maxXmlBytes, options.control);

    let hasHeaderFooter = false;
    for (const path of docx.entries.keys()) {
      if (/^word\/(header|footer)\d*\.xml$/i.test(path)) {
        hasHeaderFooter = true;
        break;
      }
    }
    report(1, "Template read");
    return {
      id: options.id,
      name: options.name,
      createdAt: options.createdAt ?? new Date().toISOString(),
      styles: [...found.values()],
      page: facts.page ?? { ...DEFAULT_PAGE },
      hasHeaderFooter,
      hasToc: facts.hasToc,
    };
  } finally {
    await docx.close().catch(() => undefined);
  }
}
