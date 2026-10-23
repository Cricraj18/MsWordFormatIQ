/**
 * Tests for zipio.ts (files 0c and 0d).
 * Run with: npm test
 * Zips are built in memory, so no sample files are needed.
 */
import { beforeAll, describe, expect, it } from "vitest";
import {
  BlobReader,
  BlobWriter,
  TextReader,
  Uint8ArrayReader,
  ZipReader,
  ZipWriter,
} from "@zip.js/zip.js";
import {
  createBlobZipOutput,
  createZipOutput,
  createJobControl,
  hasAllowedExtension,
  isSafeZipPath,
  openDocx,
  readEntryText,
  setWorkersEnabled,
  streamEntry,
} from "./zipio";
import { FormatIQError, createDefaultFormatOptions } from "../types";

beforeAll(() => {
  setWorkersEnabled(false);
});

/** Build a zip Blob from simple files. Uses raw zip.js, so paths are not checked. */
async function makeZip(files: Record<string, string | Uint8Array>): Promise<Blob> {
  const writer = new ZipWriter(new BlobWriter());
  for (const [path, content] of Object.entries(files)) {
    if (typeof content === "string") {
      await writer.add(path, new TextReader(content));
    } else {
      await writer.add(path, new Uint8ArrayReader(content));
    }
  }
  return writer.close();
}

const GOOD_FILES: Record<string, string | Uint8Array> = {
  "[Content_Types].xml": "<Types/>",
  "word/document.xml": "<w:document>Hello</w:document>",
  "word/media/image1.png": new Uint8Array([1, 2, 3, 4, 5]),
};

async function expectCode(promise: Promise<unknown>, code: string): Promise<void> {
  try {
    await promise;
  } catch (err) {
    expect(err).toBeInstanceOf(FormatIQError);
    if (err instanceof FormatIQError) expect(err.code).toBe(code);
    return;
  }
  throw new Error(`Expected error ${code} but nothing was thrown`);
}

describe("small checks", () => {
  it("accepts only .docx and .docm", () => {
    expect(hasAllowedExtension("a.docx")).toBe(true);
    expect(hasAllowedExtension("A.DOCM")).toBe(true);
    expect(hasAllowedExtension("a.pdf")).toBe(false);
    expect(hasAllowedExtension("a.docx.exe")).toBe(false);
  });

  it("accepts normal zip paths", () => {
    expect(isSafeZipPath("word/document.xml")).toBe(true);
    expect(isSafeZipPath("word/media/")).toBe(true);
  });

  it("rejects unsafe zip paths", () => {
    expect(isSafeZipPath("")).toBe(false);
    expect(isSafeZipPath("../evil.xml")).toBe(false);
    expect(isSafeZipPath("word/../../evil.xml")).toBe(false);
    expect(isSafeZipPath("/etc/passwd")).toBe(false);
    expect(isSafeZipPath("C:/x.xml")).toBe(false);
    expect(isSafeZipPath("word\\document.xml")).toBe(false);
  });

  it("rejects empty segments but allows a folder slash at the end", () => {
    expect(isSafeZipPath("a//b")).toBe(false);
    expect(isSafeZipPath("word/media/")).toBe(true);
    expect(isSafeZipPath("a/./b")).toBe(false);
  });

  it("createDefaultFormatOptions returns a new parts array each call", () => {
    const a = createDefaultFormatOptions();
    a.parts.push("toc");
    const b = createDefaultFormatOptions();
    expect(b.parts).toEqual([]);
    expect(b.imageReduction).toBe("original");
    expect(a.parts).not.toBe(b.parts);
  });
});

describe("createJobControl", () => {
  it("does not fire at once for a huge time limit", async () => {
    const control = createJobControl(2 ** 31 + 10);
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(control.signal.aborted).toBe(false);
    control.dispose();
  });

  it("times out after the limit", async () => {
    const control = createJobControl(20);
    await new Promise((resolve) => setTimeout(resolve, 80));
    expect(control.signal.aborted).toBe(true);
    expect(control.timedOut()).toBe(true);
    control.dispose();
  });
});

describe("openDocx", () => {
  it("rejects a wrong extension", async () => {
    const blob = await makeZip(GOOD_FILES);
    await expectCode(openDocx(blob, "file.pdf"), "NOT_DOCX");
  });

  it("opens a good file and lists entries", async () => {
    const opened = await openDocx(await makeZip(GOOD_FILES), "ok.docx");
    expect(opened.entries.size).toBe(3);
    expect(opened.entries.has("word/document.xml")).toBe(true);
    await opened.close();
  });

  it("rejects a zip without word/document.xml", async () => {
    const blob = await makeZip({ "[Content_Types].xml": "<Types/>" });
    await expectCode(openDocx(blob, "x.docx"), "NOT_DOCX");
  });

  it("rejects an unsafe path inside the zip", async () => {
    const blob = await makeZip({ ...GOOD_FILES, "../evil.xml": "x" });
    await expectCode(openDocx(blob, "x.docx"), "UNSAFE_ZIP_PATH");
  });

  it("rejects a document.xml over the size limit", async () => {
    const blob = await makeZip(GOOD_FILES);
    const limits = { maxEntries: 100, maxTotalBytes: 1e9, maxXmlBytes: 5, timeoutMs: 5000 };
    await expectCode(openDocx(blob, "x.docx", limits), "XML_TOO_LARGE");
  });

  it("reports CANCELLED when the job is already aborted", async () => {
    const blob = await makeZip(GOOD_FILES);
    const user = new AbortController();
    user.abort();
    const control = createJobControl(5000, user.signal);
    await expectCode(openDocx(blob, "x.docx", undefined, control), "CANCELLED");
    control.dispose();
  });

  it("reports TIMEOUT when the time limit is already over", async () => {
    const blob = await makeZip(GOOD_FILES);
    const control = createJobControl(1);
    await new Promise((resolve) => setTimeout(resolve, 30));
    await expectCode(openDocx(blob, "x.docx", undefined, control), "TIMEOUT");
    control.dispose();
  });

  it("still opens fine with a live control", async () => {
    const blob = await makeZip(GOOD_FILES);
    const control = createJobControl(5000);
    const opened = await openDocx(blob, "x.docx", undefined, control);
    expect(opened.entries.size).toBe(3);
    await opened.close();
    control.dispose();
  });
});

describe("readEntryText", () => {
  it("reads text", async () => {
    const opened = await openDocx(await makeZip(GOOD_FILES), "ok.docx");
    const entry = opened.entries.get("word/document.xml");
    expect(entry).toBeDefined();
    if (entry === undefined) return;
    expect(await readEntryText(entry, 1000)).toBe("<w:document>Hello</w:document>");
    await opened.close();
  });

  it("stops when real output is bigger than the cap", async () => {
    const opened = await openDocx(await makeZip(GOOD_FILES), "ok.docx");
    const entry = opened.entries.get("word/document.xml");
    if (entry === undefined) throw new Error("missing entry");
    await expectCode(readEntryText(entry, 5), "FILE_TOO_LARGE");
    await opened.close();
  });

  it("reports CANCELLED when the job is already aborted", async () => {
    const opened = await openDocx(await makeZip(GOOD_FILES), "ok.docx");
    const entry = opened.entries.get("word/document.xml");
    if (entry === undefined) throw new Error("missing entry");
    const user = new AbortController();
    user.abort();
    const control = createJobControl(5000, user.signal);
    await expectCode(readEntryText(entry, 1000, control), "CANCELLED");
    control.dispose();
    await opened.close();
  });
});

describe("writing", () => {
  it("copies parts untouched and adds a changed part", async () => {
    const source = await openDocx(await makeZip(GOOD_FILES), "in.docx");
    const { output, getBlob } = createBlobZipOutput();
    await output.addText("word/document.xml", "<w:document>Changed</w:document>");
    const types = source.entries.get("[Content_Types].xml");
    const image = source.entries.get("word/media/image1.png");
    if (types === undefined || image === undefined) throw new Error("missing entry");
    await output.copyEntry(types);
    await output.copyEntry(image);
    await output.finish();
    await source.close();

    const result = await openDocx(await getBlob(), "out.docx");
    const doc = result.entries.get("word/document.xml");
    if (doc === undefined) throw new Error("missing document.xml");
    expect(await readEntryText(doc, 1000)).toBe("<w:document>Changed</w:document>");
    const copied = result.entries.get("word/media/image1.png");
    if (copied === undefined) throw new Error("missing image");
    expect(await readEntryText(copied, 1000)).toBe("\u0001\u0002\u0003\u0004\u0005");
    await result.close();
  });

  it("copies a large binary part byte for byte", async () => {
    // 3 MB of repeatable pseudo-random bytes (does not compress well).
    const big = new Uint8Array(3 * 1024 * 1024);
    let x = 12345;
    for (let i = 0; i < big.length; i += 1) {
      x = (Math.imul(x, 1103515245) + 12345) & 0x7fffffff;
      big[i] = x & 0xff;
    }
    const source = await openDocx(
      await makeZip({ ...GOOD_FILES, "word/media/big.jpg": big }),
      "in.docx",
    );
    const { output, getBlob } = createBlobZipOutput();
    const entry = source.entries.get("word/media/big.jpg");
    if (entry === undefined) throw new Error("missing entry");
    await output.copyEntry(entry);
    await output.finish();
    await source.close();

    const copiedZip = await getBlob();
    const reader = new ZipReader(new BlobReader(copiedZip));
    const entries = await reader.getEntries();
    const first = entries[0];
    expect(first?.filename).toBe("word/media/big.jpg");
    expect(first?.uncompressedSize).toBe(big.length);
    if (first === undefined || first.directory) throw new Error("bad entry");
    // Compare every byte directly.
    let offset = 0;
    await streamEntry(
      first,
      (chunk) => {
        for (let i = 0; i < chunk.length; i += 1) {
          if (chunk[i] !== big[offset + i]) {
            throw new Error("Byte " + (offset + i) + " differs");
          }
        }
        offset += chunk.length;
      },
      big.length + 10,
    );
    expect(offset).toBe(big.length);
    await reader.close();
  }, 30000);

  it("writes into a WritableStream (streamed output)", async () => {
    const chunks: Uint8Array[] = [];
    const target = new WritableStream<Uint8Array>({
      write(chunk): void {
        chunks.push(chunk.slice());
      },
    });
    const progress: number[] = [];
    const output = createZipOutput(target, (info) => progress.push(info.fraction));
    await output.addText("[Content_Types].xml", "<Types/>");
    await output.addText("word/document.xml", "<w:document>Streamed</w:document>");
    await output.finish();
    expect(chunks.length).toBeGreaterThan(0);
    const result = await openDocx(new Blob(chunks.map((c) => new Uint8Array(c))), "streamed.docx");
    const doc = result.entries.get("word/document.xml");
    if (doc === undefined) throw new Error("missing document.xml");
    expect(await readEntryText(doc, 1000)).toBe("<w:document>Streamed</w:document>");
    expect(progress.length).toBeGreaterThan(0);
    await result.close();
  });

  it("refuses to write an unsafe path", async () => {
    const { output } = createBlobZipOutput();
    await expectCode(output.addText("../evil.xml", "x"), "UNSAFE_ZIP_PATH");
  });
});

describe("addBlob (version 2)", () => {
  it("writes a part from a Blob made of many text pieces", async () => {
    const pieces: string[] = ["<w:document>", "caf\u00e9 ", "\u65e5\u672c ", "</w:document>"];
    const { output, getBlob } = createBlobZipOutput();
    await output.addText("[Content_Types].xml", "<Types/>");
    await output.addBlob("word/document.xml", new Blob(pieces));
    await output.finish();
    const result = await openDocx(await getBlob(), "out.docx");
    const doc = result.entries.get("word/document.xml");
    if (doc === undefined) throw new Error("missing document.xml");
    expect(await readEntryText(doc, 1000)).toBe(pieces.join(""));
    await result.close();
  });

  it("writes a 20 MB part from a Blob and reads it back with the same size", async () => {
    const line = "<w:p><w:r><w:t>Line of text for size</w:t></w:r></w:p>";
    const count = Math.ceil((20 * 1024 * 1024) / line.length);
    const pieces: string[] = [];
    for (let i = 0; i < count; i += 1) pieces.push(line);
    const blob = new Blob(pieces);
    const { output, getBlob } = createBlobZipOutput();
    await output.addText("[Content_Types].xml", "<Types/>");
    await output.addBlob("word/document.xml", blob);
    await output.finish();
    const result = await openDocx(await getBlob(), "out.docx", { maxEntries: 100, maxTotalBytes: 1e9, maxXmlBytes: 100 * 1024 * 1024, timeoutMs: 60000 });
    const doc = result.entries.get("word/document.xml");
    if (doc === undefined) throw new Error("missing document.xml");
    expect(doc.uncompressedSize).toBe(blob.size);
    let seen = 0;
    await streamEntry(doc, (chunk) => { seen += chunk.byteLength; }, blob.size + 10);
    expect(seen).toBe(blob.size);
    await result.close();
  }, 60000);

  it("stores an already compressed type without squeezing and keeps the bytes", async () => {
    const bytes = new Uint8Array([9, 8, 7, 6, 5, 4]);
    const { output, getBlob } = createBlobZipOutput();
    await output.addText("[Content_Types].xml", "<Types/>");
    await output.addText("word/document.xml", "<w:document/>");
    await output.addBlob("word/media/pic.png", new Blob([bytes]));
    await output.finish();
    const result = await openDocx(await getBlob(), "out.docx");
    const pic = result.entries.get("word/media/pic.png");
    if (pic === undefined) throw new Error("missing image");
    const got: number[] = [];
    await streamEntry(pic, (chunk) => { got.push(...chunk); }, 100);
    expect(got).toEqual([9, 8, 7, 6, 5, 4]);
    await result.close();
  });

  it("refuses an unsafe path", async () => {
    const { output } = createBlobZipOutput();
    await expectCode(output.addBlob("../evil.xml", new Blob(["x"])), "UNSAFE_ZIP_PATH");
    await expectCode(output.addBlob("a\\b.xml", new Blob(["x"])), "UNSAFE_ZIP_PATH");
  });

  it("reports CANCELLED when the job is already aborted", async () => {
    const user = new AbortController();
    user.abort();
    const control = createJobControl(5000, user.signal);
    const { output } = createBlobZipOutput();
    await expectCode(output.addBlob("word/document.xml", new Blob(["<a/>"]), control), "CANCELLED");
    control.dispose();
  });

  it("works through a WritableStream target too", async () => {
    const chunks: Uint8Array[] = [];
    const target = new WritableStream<Uint8Array>({ write(chunk): void { chunks.push(chunk.slice()); } });
    const output = createZipOutput(target);
    await output.addText("[Content_Types].xml", "<Types/>");
    await output.addBlob("word/document.xml", new Blob(["<w:document>Blob</w:document>"]));
    await output.finish();
    const result = await openDocx(new Blob(chunks.map((c) => new Uint8Array(c))), "s.docx");
    const doc = result.entries.get("word/document.xml");
    if (doc === undefined) throw new Error("missing document.xml");
    expect(await readEntryText(doc, 1000)).toBe("<w:document>Blob</w:document>");
    await result.close();
  });
});
