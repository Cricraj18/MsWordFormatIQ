/**
 * formatIQ apply (file 3b-4, version 2: page setup and watermark wired in).
 *
 * Pours the SOURCE document into the TEMPLATE package. Output zip:
 * - from the TEMPLATE: _rels/.rels, docProps, styles, theme, fontTable, settings, webSettings,
 *   headers and footers (+ rels and media), numbering (merged);
 * - word/document.xml of the SOURCE, rewritten: template paragraph styles (detected roles),
 *   direct formatting cleared (option), list ids remapped, header/footer references replaced by
 *   the template's, page size / orientation / margins set by options.format.pageSetup;
 * - the SOURCE parts reached from its document relationships, renamed when a name clashes;
 * - a new [Content_Types].xml and word/_rels/document.xml.rels.
 * Watermark (options.format.watermark) is applied to every header part written from the template.
 * Source settings, styles, theme, fonts, headers, footers and docProps are NOT carried over.
 * The source file is never changed. Scope "selected" (3c) is not built: INVALID_OPTIONS.
 * Runs in a browser worker or Node. Never touches the network.
 */
import { FormatIQError, createDefaultFormatOptions, toFormatIQError } from "../types";
import type { DetectionResult, FormatOptions, PageSetupOptions, ProgressCallback, WatermarkOptions } from "../types";
import { DEFAULT_ZIP_LIMITS, createBlobZipOutput, createJobControl, openDocx, readEntryText, streamEntry } from "./zipio";
import type { JobControl, OpenedDocx, ZipLimits, ZipOutput } from "./zipio";
import { detectStructureInDocx } from "./structure_detect";
import { extractTemplateInfo } from "./template_extract";
import { buildStyleMap, missingStyleNames, styleIdForParagraph } from "./styles";
import { createRestyler } from "./restyle";
import type { DirectFormatting } from "./restyle";
import { createXmlRewriter } from "./xml_rewrite";
import type { RawTransform } from "./xml_rewrite";
import { chainTransforms, createNumIdRemapper, mergeNumbering, remapNumIds } from "./numbering_merge";
import {
  collectPartClosure, isTemplateOwned, mergeContentTypes, mergeRelationships, parseContentTypes,
  parseRelationships, planPartNames, pruneOverrides, relativeTarget, relsPathOf, resolveTarget,
  retargetRels, writeContentTypes, writeRelationships,
} from "./package_merge";
import type { Relationship } from "./package_merge";
import { addUpdateFields, createSectionFixer, createSectionModelReader, emptySectionModel } from "./apply_sections";
import type { SectionModel } from "./apply_sections";
import { createDefaultPageSetup, createPageSetupTransform, portraitSize, scanSections } from "./page_setup";
import { addImageRelationship, setWatermark } from "./watermark";

const MAIN_PART = "word/document.xml";
const NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const MAIN_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
/** Source parts that can hold list references (w:numId). */
const LIST_PART_PATTERN = /^word\/(footnotes|endnotes|comments)[^/]*\.xml$/i;
/** Header parts of the template. */
const HEADER_PART_PATTERN = /^word\/header[^/]*\.xml$/i;
/** Picture types allowed for an image watermark: extension -> content type. */
const IMAGE_TYPES: Readonly<Record<string, string>> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
};
/** Pieces of output text are gathered into Blobs of about this many characters. */
const FLUSH_CHARS = 1 << 20;

/** Picture for an image watermark (the bytes come from the user). */
export interface WatermarkImage {
  data: Blob;
  /** File extension without dot: png, jpg, jpeg, gif or bmp. */
  extension: string;
}

/** Options for applyTemplate. All optional. */
export interface ApplyOptions {
  /** What to apply. Default: createDefaultFormatOptions(). */
  format?: FormatOptions;
  /** "clear" (default) removes direct formatting that hides the new style. */
  directFormatting?: DirectFormatting;
  /** Detection result to use (with manual tags). Default: detect inside apply. */
  detection?: DetectionResult;
  /** Needed when format.watermark.mode is "image". */
  watermarkImage?: WatermarkImage;
  limits?: ZipLimits;
  signal?: AbortSignal;
  onProgress?: ProgressCallback;
}

/** What happened, in numbers and plain sentences. */
export interface ApplyReport {
  paragraphs: number;
  restyled: number;
  clearedParagraphProps: number;
  clearedRunProps: number;
  numIdsRemapped: number;
  numIdsUnmapped: number;
  sectionsFixed: number;
  sourceHeaderFooterRefsRemoved: number;
  /** Sections whose page size, direction or margins were written. */
  pageSetupSectionsChanged: number;
  /** Watermarks removed from template headers. */
  watermarksRemoved: number;
  /** Watermarks added to template headers. */
  watermarksAdded: number;
  sourcePartsCopied: number;
  /** [original name, output name] for source parts that were renamed. */
  renamedParts: [string, string][];
  /** Source relationships left out because the template owns that kind of part. */
  droppedSourceRelationships: number;
  /** Style names the document needs but the template does not have. */
  missingStyles: string[];
  /** True when settings.xml was changed to update fields on open. */
  fieldsUpdateOnOpen: boolean;
  warnings: string[];
}

/** Result of applyTemplate. */
export interface ApplyResult {
  blob: Blob;
  report: ApplyReport;
}

function checkOptions(format: FormatOptions): void {
  if (format.scope !== "whole") {
    throw new FormatIQError("INVALID_OPTIONS", "Applying only selected parts is not available yet. Choose the whole template.");
  }
}

/** Read the rels files of a package: owner part name -> relationships (document.xml is skipped). */
async function loadRelsByPart(docx: OpenedDocx, limits: ZipLimits, control: JobControl): Promise<Map<string, Relationship[]>> {
  const map = new Map<string, Relationship[]>();
  for (const [name, entry] of docx.entries) {
    const m = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(name);
    if (m === null) continue;
    const owner = (m[1] ?? "") + (m[2] ?? "");
    if (owner === MAIN_PART) continue;
    map.set(owner, parseRelationships(await readEntryText(entry, limits.maxXmlBytes, control), name));
  }
  return map;
}

async function readTextOf(docx: OpenedDocx, name: string, limits: ZipLimits, control: JobControl): Promise<string | undefined> {
  const entry = docx.entries.get(name);
  return entry === undefined ? undefined : readEntryText(entry, limits.maxXmlBytes, control);
}

function relOfType(rels: readonly Relationship[], suffix: string): Relationship | undefined {
  return rels.find((r) => r.type.endsWith(suffix) && r.targetMode !== "External");
}

/** Gather text pieces into Blobs so a big part never sits in memory as one string. */
function createBlobCollector(): { push: (chunk: string) => void; finish: () => Blob } {
  const blobs: Blob[] = [];
  let buffer: string[] = [];
  let size = 0;
  const flush = (): void => {
    if (buffer.length === 0) return;
    blobs.push(new Blob([buffer.join("")]));
    buffer = [];
    size = 0;
  };
  return {
    push(chunk) {
      buffer.push(chunk);
      size += chunk.length;
      if (size >= FLUSH_CHARS) flush();
    },
    finish() {
      flush();
      return new Blob(blobs);
    },
  };
}

/** Copy one part under an output name (rename allowed). */
async function copyAs(docx: OpenedDocx, name: string, outName: string, output: ZipOutput, limits: ZipLimits, control: JobControl): Promise<boolean> {
  const entry = docx.entries.get(name);
  if (entry === undefined) return false;
  if (outName === name) {
    await output.copyEntry(entry, control);
    return true;
  }
  const chunks: Uint8Array<ArrayBuffer>[] = [];
  await streamEntry(entry, (chunk) => { chunks.push(new Uint8Array(chunk)); }, limits.maxTotalBytes, control);
  await output.addBlob(outName, new Blob(chunks), control);
  return true;
}

function unique(names: readonly string[]): string[] {
  return [...new Set(names)];
}

/**
 * Apply the template to the source and write the zip into `output`.
 * Calls output.finish() at the end. Returns the report.
 */
export async function applyTemplateTo(
  output: ZipOutput,
  source: Blob,
  sourceName: string,
  template: Blob,
  templateName: string,
  options: ApplyOptions = {},
): Promise<ApplyReport> {
  const limits = options.limits ?? DEFAULT_ZIP_LIMITS;
  const format = options.format ?? createDefaultFormatOptions();
  checkOptions(format);
  const pageSetup: PageSetupOptions = format.pageSetup ?? createDefaultPageSetup();
  const watermark: WatermarkOptions = format.watermark ?? { mode: "inherit" };
  const directFormatting = options.directFormatting ?? "clear";
  const step = (phase: "reading" | "detecting" | "formatting" | "writing" | "done", fraction: number, message: string): void => {
    options.onProgress?.({ phase, fraction, message });
  };
  const control = createJobControl(limits.timeoutMs, options.signal);
  const warnings: string[] = [];
  let tpl: OpenedDocx | undefined;
  let src: OpenedDocx | undefined;
  try {
    step("reading", 0, "Opening template and document");
    tpl = await openDocx(template, templateName, limits, control);
    src = await openDocx(source, sourceName, limits, control);
    const info = await extractTemplateInfo(template, templateName, { id: "template", name: templateName, limits, control });

    // ---- watermark picture (checked early) ----
    let watermarkExt = "";
    if (watermark.mode === "image") {
      const image = options.watermarkImage;
      watermarkExt = (image?.extension ?? "").toLowerCase().replace(/^\./, "");
      if (image === undefined || IMAGE_TYPES[watermarkExt] === undefined) {
        throw new FormatIQError("INVALID_OPTIONS", "Choose a picture (png, jpg, gif or bmp) for the image watermark.");
      }
    }

    // ---- detection -> one style id per paragraph ----
    const styleMap = buildStyleMap(info.styles);
    const styleIds: (string | null)[] = [];
    const missing = new Set<string>();
    const take = (item: Parameters<typeof styleIdForParagraph>[1]): void => {
      styleIds.push(styleIdForParagraph(styleMap, item));
      for (const name of missingStyleNames(styleMap, [item])) missing.add(name);
    };
    if (options.detection !== undefined) {
      for (const item of options.detection.paragraphs) take(item);
    } else {
      await detectStructureInDocx(src, take, {
        limits, control,
        onProgress: (p) => step("detecting", p.fraction * 0.3, p.message),
      });
    }

    // ---- relationships ----
    step("reading", 0.35, "Reading relationships");
    const tplRootXml = await readTextOf(tpl, "_rels/.rels", limits, control);
    if (tplRootXml === undefined) throw new FormatIQError("TEMPLATE_INVALID", "The template has no _rels/.rels part.");
    const tplDocRelsRaw = parseRelationships(await readTextOf(tpl, "word/_rels/document.xml.rels", limits, control) ?? "<Relationships/>", "template document.xml.rels");
    const srcDocRels = parseRelationships(await readTextOf(src, "word/_rels/document.xml.rels", limits, control) ?? "<Relationships/>", "source document.xml.rels");
    const tplRelsByPart = await loadRelsByPart(tpl, limits, control);
    const srcRelsByPart = await loadRelsByPart(src, limits, control);

    const tplDocRels = tplDocRelsRaw.filter((r) => {
      if (!isTemplateOwned(r.type) || r.targetMode === "External") return true;
      const target = resolveTarget("word", r.target);
      const ok = target !== undefined && tpl !== undefined && tpl.entries.has(target);
      if (!ok) warnings.push(`Template relationship ${r.id} points to a missing part (${r.target}) and was left out.`);
      return ok;
    });
    const mergedRels = mergeRelationships(srcDocRels, tplDocRels);

    // ---- which parts are copied ----
    const rootRels = parseRelationships(tplRootXml, "template _rels/.rels");
    const tplRootClosure = collectPartClosure(rootRels, "", tplRelsByPart).filter((n) => n !== MAIN_PART);
    const tplOwnedClosure = collectPartClosure(tplDocRels.filter((r) => isTemplateOwned(r.type)), "word", tplRelsByPart);
    const tplParts = unique([...tplRootClosure, ...tplOwnedClosure]).filter((n) => {
      const there = tpl !== undefined && tpl.entries.has(n);
      if (!there) warnings.push(`Template part ${n} is named in a relationship but is missing; skipped.`);
      return there;
    });
    const srcKept = srcDocRels.filter((r) => !isTemplateOwned(r.type));
    const srcParts = unique(collectPartClosure(srcKept, "word", srcRelsByPart)).filter((n) => {
      if (n === MAIN_PART) return false;
      const there = src !== undefined && src.entries.has(n);
      if (!there) warnings.push(`Source part ${n} is named in a relationship but is missing; skipped.`);
      return there;
    });

    // ---- numbering ----
    const tplNumRel = relOfType(tplDocRels, "/numbering");
    const tplNumPart = tplNumRel === undefined ? undefined : resolveTarget("word", tplNumRel.target);
    const srcNumRel = relOfType(srcDocRels, "/numbering");
    const srcNumPartName = srcNumRel === undefined ? undefined : resolveTarget("word", srcNumRel.target);
    const tplNumXml = tplNumPart === undefined ? undefined : await readTextOf(tpl, tplNumPart, limits, control);
    const srcNumXml = srcNumPartName === undefined ? undefined : await readTextOf(src, srcNumPartName, limits, control);
    const numbering = mergeNumbering(tplNumXml, srcNumXml, srcNumPartName ?? "word/numbering.xml");
    warnings.push(...numbering.warnings);

    const taken = new Set<string>([...tplParts, MAIN_PART]);
    let newNumberingPart: string | undefined;
    if (numbering.xml !== undefined && tplNumPart === undefined) {
      newNumberingPart = "word/numbering.xml";
      for (let n = 2; [...taken].some((t) => t.toLowerCase() === newNumberingPart?.toLowerCase()); n += 1) {
        newNumberingPart = `word/numbering${n}.xml`;
      }
      taken.add(newNumberingPart);
    }

    // ---- watermark plan: which template headers, which picture name ----
    const headerParts = tplParts.filter((n) => HEADER_PART_PATTERN.test(n)).sort();
    const watermarkActive = watermark.mode === "text" || watermark.mode === "image" || watermark.mode === "remove";
    if (watermark.mode === "keepSource") {
      warnings.push("Keeping the source watermark needs the source headers, which are not carried over yet; the template's headers and watermark are used.");
    }
    if ((watermark.mode === "text" || watermark.mode === "image") && headerParts.length === 0) {
      warnings.push("The template has no header, so the watermark was not added.");
    }
    let watermarkPicturePart: string | undefined;
    if (watermark.mode === "image" && headerParts.length > 0) {
      const wanted = "word/media/formatiq_watermark." + watermarkExt;
      const planned = planPartNames([wanted], [...taken]);
      watermarkPicturePart = planned.get(wanted) ?? wanted;
      taken.add(watermarkPicturePart);
    }
    const renames = planPartNames(srcParts, taken);

    // ---- content types and relationships of the main part ----
    const tplCt = parseContentTypes(await readTextOf(tpl, "[Content_Types].xml", limits, control) ?? "");
    const srcCt = parseContentTypes(await readTextOf(src, "[Content_Types].xml", limits, control) ?? "");
    const ct = mergeContentTypes(tplCt, srcCt, renames);
    if (!ct.defaults.has("rels")) ct.defaults.set("rels", RELS_CONTENT_TYPE);
    if (!ct.defaults.has("xml")) ct.defaults.set("xml", "application/xml");
    if (!ct.overrides.has("/" + MAIN_PART)) ct.overrides.set("/" + MAIN_PART, tplCt.overrides.get("/" + MAIN_PART) ?? MAIN_CONTENT_TYPE);
    if (newNumberingPart !== undefined) ct.overrides.set("/" + newNumberingPart, NUMBERING_CONTENT_TYPE);
    if (watermarkPicturePart !== undefined && !ct.defaults.has(watermarkExt)) {
      ct.defaults.set(watermarkExt, IMAGE_TYPES[watermarkExt] ?? "application/octet-stream");
    }
    const outputParts = [
      ...tplParts, MAIN_PART, ...(newNumberingPart === undefined ? [] : [newNumberingPart]),
      ...(watermarkPicturePart === undefined ? [] : [watermarkPicturePart]), ...renames.values(),
    ];
    const finalCt = pruneOverrides(ct, outputParts);

    const templateIds = new Set(mergedRels.templateIdMap.values());
    const sourceRelsOut = retargetRels(mergedRels.rels.filter((r) => !templateIds.has(r.id)), MAIN_PART, MAIN_PART, renames);
    const templateRelsOut = mergedRels.rels.filter((r) => templateIds.has(r.id));
    const docRelsOut: Relationship[] = [...sourceRelsOut, ...templateRelsOut];
    if (newNumberingPart !== undefined) {
      let id = "tplNumbering";
      for (let n = 2; docRelsOut.some((r) => r.id === id); n += 1) id = "tplNumbering" + n;
      docRelsOut.push({ id, type: NUMBERING_REL_TYPE, target: relativeTarget("word", newNumberingPart) });
    }

    // ---- template document.xml: header/footer model and first section ----
    let sectionModel: SectionModel = emptySectionModel();
    let templateSection = undefined as ReturnType<typeof scanSections>[number] | undefined;
    const tplMain = tpl.entries.get(MAIN_PART);
    if (tplMain !== undefined) {
      const tplMainText = await readEntryText(tplMain, limits.maxXmlBytes, control);
      const reader = createSectionModelReader();
      const feeder = createXmlRewriter({ partName: "template " + MAIN_PART, onOutput: () => undefined, transform: reader.transform });
      feeder.write(tplMainText);
      feeder.end();
      sectionModel = reader.getModel();
      templateSection = scanSections(tplMainText, "template " + MAIN_PART)[0];
    }
    portraitSize(pageSetup, templateSection);

    // ---- write: content types and root rels first ----
    step("writing", 0.4, "Writing the new document");
    await output.addText("[Content_Types].xml", writeContentTypes(finalCt), control);
    await output.addText("_rels/.rels", tplRootXml, control);

    // ---- word/document.xml ----
    const srcMain = src.entries.get(MAIN_PART);
    if (srcMain === undefined) throw new FormatIQError("NOT_DOCX", "This does not look like a Word document.");
    const remapper = createNumIdRemapper(numbering.numIdMap);
    const restyler = createRestyler({ styleFor: (i) => styleIds[i] ?? null, directFormatting });
    const fixer = createSectionFixer(sectionModel, mergedRels.templateIdMap);
    const pageTransform = createPageSetupTransform(pageSetup, templateSection);
    let sawFields = false;
    const fieldWatcher: RawTransform = (token, emit) => {
      if (token.kind === "open") {
        const colon = token.name.indexOf(":");
        const local = colon === -1 ? token.name : token.name.slice(colon + 1);
        if (local === "fldChar" || local === "fldSimple" || local === "instrText") sawFields = true;
      }
      emit(token);
    };
    const transform = chainTransforms(
      chainTransforms(remapper.transform, restyler.transform),
      chainTransforms(fixer.transform, chainTransforms(pageTransform.transform, fieldWatcher)),
    );
    const collector = createBlobCollector();
    const rewriter = createXmlRewriter({ partName: MAIN_PART, onOutput: collector.push, transform });
    const decoder = new TextDecoder("utf-8");
    let readBytes = 0;
    const total = Math.max(1, srcMain.uncompressedSize);
    await streamEntry(srcMain, (chunk) => {
      readBytes += chunk.byteLength;
      rewriter.write(decoder.decode(chunk, { stream: true }));
      step("formatting", 0.3 + 0.5 * Math.min(1, readBytes / total), "Formatting the document");
    }, limits.maxXmlBytes, control);
    const rest = decoder.decode();
    if (rest.length > 0) rewriter.write(rest);
    rewriter.end();
    await output.addBlob(MAIN_PART, collector.finish(), control);
    if (restyler.stats.paragraphs !== styleIds.length) {
      warnings.push(`The document has ${restyler.stats.paragraphs} paragraphs but ${styleIds.length} were detected; extra paragraphs were left as they are.`);
    }
    await output.addText("word/_rels/document.xml.rels", writeRelationships(docRelsOut), control);

    // ---- template parts ----
    step("writing", 0.85, "Copying template parts");
    const settingsRel = relOfType(tplDocRels, "/settings");
    const tplSettingsPart = settingsRel === undefined ? undefined : resolveTarget("word", settingsRel.target);
    let fieldsUpdateOnOpen = false;
    let watermarksRemoved = 0;
    let watermarksAdded = 0;
    for (const name of tplParts) {
      let relsHandled = false;
      if (name === tplNumPart) {
        if (numbering.xml !== undefined) await output.addText(name, numbering.xml, control);
      } else if (name === tplSettingsPart && format.updateFieldsOnOpen && sawFields) {
        const text = await readTextOf(tpl, name, limits, control);
        if (text === undefined) continue;
        const result = addUpdateFields(text, name);
        fieldsUpdateOnOpen = result.changed;
        await output.addText(name, result.xml, control);
      } else if (watermarkActive && HEADER_PART_PATTERN.test(name)) {
        const text = await readTextOf(tpl, name, limits, control);
        if (text === undefined) continue;
        const seed = headerParts.indexOf(name) + 1;
        let wm: WatermarkOptions = { ...watermark, seed };
        let relsText: string | undefined;
        if (watermark.mode === "image" && watermarkPicturePart !== undefined) {
          const relsName = relsPathOf(name);
          const oldRels = await readTextOf(tpl, relsName, limits, control);
          const added = addImageRelationship(oldRels, relativeTarget("word", watermarkPicturePart));
          relsText = added.xml;
          wm = { ...wm, imageRelId: added.id };
        }
        const result = setWatermark(text, name, wm);
        watermarksRemoved += result.removed;
        if (result.added) watermarksAdded += 1;
        await output.addText(name, result.xml, control);
        if (relsText !== undefined && result.added) {
          await output.addText(relsPathOf(name), relsText, control);
          relsHandled = true;
        }
      } else {
        await copyAs(tpl, name, name, output, limits, control);
      }
      if (!relsHandled) {
        const relsEntry = tpl.entries.get(relsPathOf(name));
        if (relsEntry !== undefined) await output.copyEntry(relsEntry, control);
      }
    }
    if (watermarkPicturePart !== undefined && options.watermarkImage !== undefined && watermarksAdded > 0) {
      await output.addBlob(watermarkPicturePart, options.watermarkImage.data, control);
    }
    if (newNumberingPart !== undefined && numbering.xml !== undefined) {
      await output.addText(newNumberingPart, numbering.xml, control);
    }
    if (format.updateFieldsOnOpen && sawFields && tplSettingsPart === undefined) {
      warnings.push("The document has fields, but the template has no settings part, so Word will not update them on open.");
    }

    // ---- source parts ----
    step("writing", 0.92, "Copying document parts");
    let numRemapped = remapper.stats.remapped;
    let numUnmapped = remapper.stats.unmapped;
    for (const original of srcParts) {
      const outName = renames.get(original) ?? original;
      if (numbering.numIdMap.size > 0 && LIST_PART_PATTERN.test(original)) {
        const text = await readTextOf(src, original, limits, control);
        if (text !== undefined) {
          const result = remapNumIds(text, original, numbering.numIdMap);
          numRemapped += result.stats.remapped;
          numUnmapped += result.stats.unmapped;
          await output.addText(outName, result.xml, control);
        }
      } else {
        await copyAs(src, original, outName, output, limits, control);
      }
      const rels = srcRelsByPart.get(original);
      if (rels !== undefined) {
        await output.addText(relsPathOf(outName), writeRelationships(retargetRels(rels, original, outName, renames)), control);
      }
    }
    await output.finish();
    step("done", 1, "Done");

    warnings.push("Source settings, styles, theme, fonts, headers, footers and document properties were not carried over; the template's are used.");
    if (numUnmapped > 0) warnings.push(`${numUnmapped} list reference(s) point to a list that is not in the numbering part and were left unchanged.`);
    if (sawFields) warnings.push("The document has fields (page numbers, references...). They may look empty until Word updates them.");
    return {
      paragraphs: restyler.stats.paragraphs,
      restyled: restyler.stats.restyled,
      clearedParagraphProps: restyler.stats.clearedParagraphProps,
      clearedRunProps: restyler.stats.clearedRunProps,
      numIdsRemapped: numRemapped,
      numIdsUnmapped: numUnmapped,
      sectionsFixed: fixer.stats.sections,
      sourceHeaderFooterRefsRemoved: fixer.stats.removedSourceRefs,
      pageSetupSectionsChanged: pageTransform.stats.changed,
      watermarksRemoved,
      watermarksAdded,
      sourcePartsCopied: srcParts.length,
      renamedParts: [...renames].filter(([a, b]) => a !== b),
      droppedSourceRelationships: mergedRels.droppedSourceRels.length,
      missingStyles: [...missing].sort(),
      fieldsUpdateOnOpen,
      warnings,
    };
  } catch (error) {
    throw toFormatIQError(error);
  } finally {
    control.dispose();
    if (tpl !== undefined) await tpl.close().catch(() => undefined);
    if (src !== undefined) await src.close().catch(() => undefined);
  }
}

/** Apply and return the result as a Blob. Fine for small and medium files. */
export async function applyTemplate(
  source: Blob,
  sourceName: string,
  template: Blob,
  templateName: string,
  options: ApplyOptions = {},
): Promise<ApplyResult> {
  const { output, getBlob } = createBlobZipOutput();
  const report = await applyTemplateTo(output, source, sourceName, template, templateName, options);
  return { blob: await getBlob(), report };
}
