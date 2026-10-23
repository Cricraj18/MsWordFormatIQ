/**
 * formatIQ shared types (file 0b, version 2: page setup, watermark, header/footer choices).
 *
 * This file holds the small set of types that many other files share.
 * It has no imports, so any file can import from it without loops.
 *
 * Rules kept here: no "any", no non-null assertions, strict typing.
 * Everything runs in the browser. Nothing here touches the network.
 */

// ---------------------------------------------------------------------------
// Progress
// ---------------------------------------------------------------------------

/** Which stage of the job is running. Used for the progress bar text. */
export type ProgressPhase =
  | "reading"
  | "detecting"
  | "formatting"
  | "images"
  | "writing"
  | "done";

/** One progress update sent from a worker to the page. */
export interface ProgressInfo {
  /** Current stage. */
  phase: ProgressPhase;
  /** Number from 0 to 1. 0 = just started, 1 = finished. */
  fraction: number;
  /** Short human message, for example "Reading document.xml". */
  message: string;
}

/** Function that receives progress updates. */
export type ProgressCallback = (info: ProgressInfo) => void;

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

/** Machine-readable reason for a formatIQ error. */
export type FormatIQErrorCode =
  | "NOT_DOCX" // file is not .docx or .docm
  | "FILE_TOO_LARGE" // unzipped size over the limit
  | "XML_TOO_LARGE" // document.xml over the limit
  | "UNSAFE_ZIP_PATH" // zip entry path tries to escape (for example "../")
  | "XML_ENTITY_EXPANSION" // XML uses entity tricks (blocked)
  | "TIMEOUT" // job took longer than the time limit
  | "CANCELLED" // user pressed Cancel
  | "ZIP_READ_FAILED" // could not read the zip
  | "ZIP_WRITE_FAILED" // could not write the output zip
  | "TEMPLATE_INVALID" // template .docx is missing something needed
  | "INVALID_OPTIONS" // FormatOptions has a bad combination
  | "UNKNOWN"; // anything else

/**
 * The one error type formatIQ code throws on purpose.
 * Catch it and show `message` to the user. Use `code` in your logic.
 */
export class FormatIQError extends Error {
  /** Machine-readable reason. */
  public readonly code: FormatIQErrorCode;
  /** Optional extra detail, for example a file path inside the zip. */
  public readonly detail: string | undefined;

  constructor(code: FormatIQErrorCode, message: string, detail?: string) {
    super(message);
    this.name = "FormatIQError";
    this.code = code;
    this.detail = detail;
    // Keeps "instanceof" working even if the build target is old.
    Object.setPrototypeOf(this, FormatIQError.prototype);
  }
}

/** True if the value is a FormatIQError. */
export function isFormatIQError(value: unknown): value is FormatIQError {
  return value instanceof FormatIQError;
}

/**
 * Turn anything that was thrown into a FormatIQError.
 * Use inside catch blocks so the rest of the code deals with one type.
 */
export function toFormatIQError(value: unknown): FormatIQError {
  if (value instanceof FormatIQError) {
    return value;
  }
  if (value instanceof Error) {
    return new FormatIQError("UNKNOWN", value.message);
  }
  return new FormatIQError("UNKNOWN", "Unknown error", String(value));
}

// ---------------------------------------------------------------------------
// Template
// ---------------------------------------------------------------------------

/** What kind of thing a Word style applies to. */
export type StyleKind = "paragraph" | "character" | "table" | "numbering";

/** One style found in the template's styles.xml. */
export interface TemplateStyle {
  /** Style id used inside the XML, for example "Heading1". */
  id: string;
  /** Name shown in Word, for example "heading 1". */
  name: string;
  kind: StyleKind;
}

/** Page setup of the template's main section. Sizes are in twips. */
export interface PageSetup {
  widthTwips: number;
  heightTwips: number;
  landscape: boolean;
}

/** Facts read from a template .docx. Produced by template_extract.ts. */
export interface TemplateInfo {
  /** Unique id used when storing the template in the browser. */
  id: string;
  /** Name shown to the user. */
  name: string;
  /** When it was added, in ISO format, for example "2026-10-11T07:00:00Z". */
  createdAt: string;
  /** All styles the template offers. */
  styles: TemplateStyle[];
  /** Page setup of the first section. */
  page: PageSetup;
  /** True if the template has any header or footer parts. */
  hasHeaderFooter: boolean;
  /** True if the template already has a table of contents field. */
  hasToc: boolean;
}

// ---------------------------------------------------------------------------
// Structure detection
// ---------------------------------------------------------------------------

/** The role a paragraph plays in the document. */
export type StructureRole =
  | "title"
  | "heading"
  | "body"
  | "list-item"
  | "figure-caption"
  | "table-caption"
  | "table"
  | "quote"
  | "code"
  | "unknown";

/** Facts about one paragraph, read from the source document. */
export interface ParagraphInfo {
  /** Position in the document, starting at 0. */
  index: number;
  /** Plain text of the paragraph. Can be empty. */
  text: string;
  /** Style id the paragraph has now, or null if none. */
  styleId: string | null;
  /** True if the paragraph sits inside a table cell. */
  inTable: boolean;
  /** True if the paragraph contains a picture. */
  hasImage: boolean;
  /** True if the paragraph is part of a numbered or bulleted list. */
  isListItem: boolean;
}

/** Detection result for one paragraph. */
export interface DetectedParagraph {
  paragraph: ParagraphInfo;
  /** Guessed role. */
  role: StructureRole;
  /** Heading level 1 to 9 when role is "heading", otherwise null. */
  level: number | null;
  /** Confidence from 0 to 1. Low values need manual tagging. */
  confidence: number;
  /** Role the user chose by hand, or null if not tagged. Wins over `role`. */
  manualRole: StructureRole | null;
}

/** Below this confidence, the user is asked to tag the paragraph by hand. */
export const LOW_CONFIDENCE_THRESHOLD = 0.6;

/** Output of structure_detect.ts. */
export interface DetectionResult {
  paragraphs: DetectedParagraph[];
  /** Number of paragraphs below LOW_CONFIDENCE_THRESHOLD. */
  lowConfidenceCount: number;
}

// ---------------------------------------------------------------------------
// Page setup (moved here from page_setup.ts, HANDOFF v24 TYPES NOTE)
// ---------------------------------------------------------------------------

/** Paper size choice. "template" keeps the template's own size. */
export type PagePreset =
  | "template"
  | "a4"
  | "letter"
  | "a3"
  | "legal"
  | "a5"
  | "tabloid"
  | "custom";

/** Page direction. "template" follows the template (a landscape source section stays landscape). */
export type Orientation = "template" | "portrait" | "landscape";

/** Page margins and header/footer distances, all in twips. */
export interface MarginSet {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter: number;
}

/** What the user chose for page size, direction and margins. */
export interface PageSetupOptions {
  preset: PagePreset;
  /** Used only when preset is "custom" (twips). */
  customWidthTwips?: number;
  /** Used only when preset is "custom" (twips). */
  customHeightTwips?: number;
  orientation: Orientation;
  /** "template" keeps the template margins. */
  margins: MarginSet | "template";
  /** "all" sections, or a list of section numbers (starting at 0). */
  sections: "all" | number[];
}

// ---------------------------------------------------------------------------
// Watermark (moved here from watermark.ts, HANDOFF v26)
// ---------------------------------------------------------------------------

/**
 * What to do with watermarks.
 * "inherit" leaves the template's watermark as it is (default).
 * "remove" removes it. "keepSource" keeps the source's (needs keepSourceHeaders).
 * "text" and "image" replace it with a new one.
 */
export type WatermarkMode = "inherit" | "remove" | "keepSource" | "text" | "image";

/** Watermark choice. Only `mode` is required. */
export interface WatermarkOptions {
  mode: WatermarkMode;
  /** Text watermark: the words to show. */
  text?: string;
  /** Colour as RRGGBB, default C0C0C0. */
  color?: string;
  /** More than 0, at most 1, default 0.5. */
  opacity?: number;
  /** Diagonal (default true) or horizontal. */
  diagonal?: boolean;
  /** Font name, default Calibri. */
  fontFamily?: string;
  /** Image watermark: relationship id of the picture in the header. */
  imageRelId?: string;
  /** Image watermark: size in points (both needed, up to 1500). */
  widthPt?: number;
  heightPt?: number;
  /** Image watermark: washed-out look, default true. */
  washout?: boolean;
  /** Whole number 1..1000000 that makes shape ids differ per header, default 1. */
  seed?: number;
}

// ---------------------------------------------------------------------------
// Format options
// ---------------------------------------------------------------------------

/** Image reduction level. "original" means no change (the default). */
export type ImageReductionLevel = "original" | "light" | "balanced";

/** Parts of the template that can be applied on their own. */
export type FormatPart =
  | "styles"
  | "headersFooters"
  | "toc"
  | "crossReferences"
  | "sections";

/** What the user chose before pressing Format. */
export interface FormatOptions {
  /** "whole" applies everything. "selected" applies only `parts`. */
  scope: "whole" | "selected";
  /** Used only when scope is "selected". */
  parts: FormatPart[];
  /** Add table of contents, list of figures, list of tables. */
  includeToc: boolean;
  includeListOfFigures: boolean;
  includeListOfTables: boolean;
  /** Mark the file so Word updates fields when it opens. */
  updateFieldsOnOpen: boolean;
  /** Image reduction. Keep "original" unless the user turns it on. */
  imageReduction: ImageReductionLevel;
  /**
   * Page size, direction, margins. Missing = everything follows the template.
   * OPTIONAL on purpose: older code and tests build FormatOptions by hand without it.
   * createDefaultFormatOptions() always fills it. Make it required after every importer is updated.
   */
  pageSetup?: PageSetupOptions;
  /** Watermark choice. Missing = "inherit". Optional for the same reason as pageSetup. */
  watermark?: WatermarkOptions;
  /** Keep the source document's own headers and footers. Missing = false (the template's are used). */
  keepSourceHeaders?: boolean;
}

/**
 * Safe starting values. Image reduction is OFF. Page setup follows the template,
 * watermark is "inherit", source headers are dropped.
 * Returns a NEW object (and new nested objects and arrays) on every call, so callers
 * can change it without touching anyone else's copy.
 */
export function createDefaultFormatOptions(): FormatOptions {
  return {
    scope: "whole",
    parts: [],
    includeToc: true,
    includeListOfFigures: false,
    includeListOfTables: false,
    updateFieldsOnOpen: true,
    imageReduction: "original",
    pageSetup: {
      preset: "template",
      orientation: "template",
      margins: "template",
      sections: "all",
    },
    watermark: { mode: "inherit" },
    keepSourceHeaders: false,
  };
}
