/**
 * Tests for types.ts (version 2). Run with: npm test
 */
import { describe, expect, it } from "vitest";
import {
  FormatIQError,
  LOW_CONFIDENCE_THRESHOLD,
  createDefaultFormatOptions,
  isFormatIQError,
  toFormatIQError,
} from "./types";
import type { FormatOptions, MarginSet, PageSetupOptions, WatermarkOptions } from "./types";

describe("createDefaultFormatOptions", () => {
  it("keeps the old defaults", () => {
    const o = createDefaultFormatOptions();
    expect(o.scope).toBe("whole");
    expect(o.parts).toEqual([]);
    expect(o.includeToc).toBe(true);
    expect(o.includeListOfFigures).toBe(false);
    expect(o.includeListOfTables).toBe(false);
    expect(o.updateFieldsOnOpen).toBe(true);
    expect(o.imageReduction).toBe("original");
  });

  it("page setup follows the template, all sections", () => {
    const o = createDefaultFormatOptions();
    expect(o.pageSetup).toEqual({ preset: "template", orientation: "template", margins: "template", sections: "all" });
  });

  it("watermark is inherit and source headers are dropped", () => {
    const o = createDefaultFormatOptions();
    expect(o.watermark).toEqual({ mode: "inherit" });
    expect(o.keepSourceHeaders).toBe(false);
  });

  it("returns new nested objects on every call", () => {
    const a = createDefaultFormatOptions();
    a.parts.push("toc");
    if (a.pageSetup !== undefined) a.pageSetup.preset = "a4";
    if (a.watermark !== undefined) a.watermark.mode = "remove";
    const b = createDefaultFormatOptions();
    expect(b.parts).toEqual([]);
    expect(b.pageSetup?.preset).toBe("template");
    expect(b.watermark?.mode).toBe("inherit");
    expect(a.pageSetup).not.toBe(b.pageSetup);
    expect(a.watermark).not.toBe(b.watermark);
  });
});

describe("old hand-made options still compile", () => {
  it("FormatOptions without the new fields is allowed", () => {
    const old: FormatOptions = {
      scope: "selected", parts: [], includeToc: false, includeListOfFigures: false,
      includeListOfTables: false, updateFieldsOnOpen: false, imageReduction: "original",
    };
    expect(old.pageSetup).toBeUndefined();
    expect(old.watermark).toBeUndefined();
    expect(old.keepSourceHeaders).toBeUndefined();
  });
});

describe("page setup and watermark types", () => {
  it("accept full values", () => {
    const margins: MarginSet = { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 };
    const ps: PageSetupOptions = { preset: "custom", customWidthTwips: 10000, customHeightTwips: 14000, orientation: "landscape", margins, sections: [0, 2] };
    const wm: WatermarkOptions = { mode: "text", text: "DRAFT", color: "C0C0C0", opacity: 0.5, diagonal: true, fontFamily: "Calibri", seed: 3 };
    const img: WatermarkOptions = { mode: "image", imageRelId: "rIdWm1", widthPt: 300, heightPt: 200, washout: true };
    expect(ps.sections).toEqual([0, 2]);
    expect(wm.text).toBe("DRAFT");
    expect(img.imageRelId).toBe("rIdWm1");
  });
});

describe("errors and constants (unchanged behaviour)", () => {
  it("FormatIQError keeps code and detail", () => {
    const e = new FormatIQError("INVALID_OPTIONS", "bad", "section 2");
    expect(isFormatIQError(e)).toBe(true);
    expect(e instanceof Error).toBe(true);
    expect(e.code).toBe("INVALID_OPTIONS");
    expect(e.detail).toBe("section 2");
    expect(e.name).toBe("FormatIQError");
  });

  it("toFormatIQError wraps other values", () => {
    expect(toFormatIQError(new Error("x")).code).toBe("UNKNOWN");
    expect(toFormatIQError("oops").detail).toBe("oops");
    const own = new FormatIQError("CANCELLED", "c");
    expect(toFormatIQError(own)).toBe(own);
  });

  it("low confidence threshold is 0.6", () => {
    expect(LOW_CONFIDENCE_THRESHOLD).toBe(0.6);
  });
});
