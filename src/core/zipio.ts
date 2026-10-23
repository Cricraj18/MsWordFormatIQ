/**
 * formatIQ zip input and output (files 0c and 0d).
 *
 * 0c = reading: open a .docx lazily (only the zip index is read at first),
 * check it is safe, then read single parts as streams.
 * 0d = writing: build a new zip, copy untouched parts, add changed parts.
 *
 * Version 2 (chat 10): ZipOutput.addBlob added, so a big rewritten part
 * (for example word/document.xml) can be written from a Blob instead of one
 * huge string.
 *
 * Library: @zip.js/zip.js. Tested with zip.js 2.23.0 (Node 22 and Node 18).
 * Not tested in a real browser: Web Workers, showSaveFilePicker.
 *
 * Runs in the browser (page or Web Worker). Never touches the network.
 */
import {
  BlobReader,
  BlobWriter,
  TextReader,
  ZipReader,
  ZipWriter,
  configure,
} from "@zip.js/zip.js";
import type { Entry, FileEntry } from "@zip.js/zip.js";
import { FormatIQError, isFormatIQError } from "../types";
import type { FormatIQErrorCode, ProgressCallback } from "../types";

// ---------------------------------------------------------------------------
// Limits and small checks
// ---------------------------------------------------------------------------

/** Safety limits. All sizes are in bytes. */
export interface ZipLimits {
  /** Most entries (files) allowed in one zip. */
  maxEntries: number;
  /** Most total unzipped size allowed. */
  maxTotalBytes: number;
  /** Most size allowed for any single XML part, for example document.xml. */
  maxXmlBytes: number;
  /** Time limit for one job, in milliseconds. */
  timeoutMs: number;
}

/** Starting limits. A 600 MB file was tested; a true 1 GB file was not. */
export const DEFAULT_ZIP_LIMITS: ZipLimits = {
  maxEntries: 50_000,
  maxTotalBytes: 4 * 1024 * 1024 * 1024,
  maxXmlBytes: 256 * 1024 * 1024,
  timeoutMs: 10 * 60 * 1000,
};

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** Largest delay setTimeout accepts. Bigger values fire at once. */
const MAX_TIMER_MS = 2_147_483_647;

/** Turn zip.js web workers on or off. Tests turn them off. */
export function setWorkersEnabled(enabled: boolean): void {
  configure({ useWebWorkers: enabled });
}

/** True only for names ending in .docx or .docm (any letter case). */
export function hasAllowedExtension(fileName: string): boolean {
  return /\.(docx|docm)$/i.test(fileName);
}

/**
 * True if a path inside a zip is safe.
 * Blocks empty, absolute, backslash, drive letter, control characters,
 * empty segments ("a//b") and any "." or ".." segment.
 * One trailing "/" (a folder entry such as "word/media/") is allowed.
 */
export function isSafeZipPath(path: string): boolean {
  if (path.length === 0 || path.length > 512) return false;
  if (path.startsWith("/") || path.includes("\\")) return false;
  if (/^[A-Za-z]:/.test(path)) return false;
  for (let i = 0; i < path.length; i += 1) {
    if (path.charCodeAt(i) < 32) return false;
  }
  return path
    .replace(/\/$/, "")
    .split("/")
    .every((part) => part !== "" && part !== ".." && part !== ".");
}

// ---------------------------------------------------------------------------
// Job control: cancel button and time limit
// ---------------------------------------------------------------------------

/** Lets a job be cancelled or time-limited. Call dispose() when finished. */
export interface JobControl {
  signal: AbortSignal;
  timedOut(): boolean;
  dispose(): void;
}

/** Make a JobControl. Aborts on time limit or when userSignal aborts. */
export function createJobControl(
  timeoutMs: number,
  userSignal?: AbortSignal,
): JobControl {
  const controller = new AbortController();
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, Math.min(timeoutMs, MAX_TIMER_MS));
  const onUserAbort = (): void => controller.abort();
  if (userSignal) {
    if (userSignal.aborted) controller.abort();
    else userSignal.addEventListener("abort", onUserAbort, { once: true });
  }
  return {
    signal: controller.signal,
    timedOut: () => timedOut,
    dispose: () => {
      clearTimeout(timer);
      if (userSignal) userSignal.removeEventListener("abort", onUserAbort);
    },
  };
}

/** Turn any thrown value into a FormatIQError with a helpful code. */
function mapError(
  err: unknown,
  control: JobControl | undefined,
  fallback: FormatIQErrorCode,
): FormatIQError {
  if (isFormatIQError(err)) return err;
  if (control !== undefined && control.signal.aborted) {
    return control.timedOut()
      ? new FormatIQError("TIMEOUT", "This took too long and was stopped.")
      : new FormatIQError("CANCELLED", "Cancelled.");
  }
  const text = err instanceof Error ? err.message : String(err);
  // zip.js refuses unsafe names itself while reading. Its message is
  // "Unsafe filename" (tested with zip.js 2.23.0; could change later).
  if (/unsafe filename/i.test(text)) {
    return new FormatIQError("UNSAFE_ZIP_PATH", "The file has an unsafe path.", text);
  }
  return new FormatIQError(fallback, "The zip file could not be processed.", text);
}

function signalOptions(
  control: JobControl | undefined,
): { signal: AbortSignal } | undefined {
  return control === undefined ? undefined : { signal: control.signal };
}

// ---------------------------------------------------------------------------
// 0c: reading
// ---------------------------------------------------------------------------

/** An opened, checked .docx. Close it when finished. */
export interface OpenedDocx {
  /** All file entries by path. Folders are left out. */
  entries: Map<string, FileEntry>;
  /** Sum of unzipped sizes as written in the zip index. */
  totalBytes: number;
  /** Release the reader. */
  close(): Promise<void>;
}

/**
 * Open a .docx or .docm without loading it all into memory.
 * Checks: extension, entry count, safe paths, no duplicates, no passwords,
 * total size, document.xml size, and that the two core parts exist.
 *
 * getEntries() itself cannot be cancelled (it only reads the zip index, so
 * it is short). The control is checked before and after it.
 *
 * @param source Usually a File from the file picker (a File is a Blob).
 * @param fileName Original file name, used for the extension check.
 * @param limits Safety limits.
 * @param control Optional cancel button and time limit.
 */
export async function openDocx(
  source: Blob,
  fileName: string,
  limits: ZipLimits = DEFAULT_ZIP_LIMITS,
  control?: JobControl,
): Promise<OpenedDocx> {
  if (!hasAllowedExtension(fileName)) {
    throw new FormatIQError("NOT_DOCX", "Please choose a .docx or .docm file.", fileName);
  }
  const reader = new ZipReader(new BlobReader(source));
  try {
    if (control !== undefined && control.signal.aborted) {
      throw new Error("aborted before reading the zip index");
    }
    const all: Entry[] = await reader.getEntries();
    if (control !== undefined && control.signal.aborted) {
      throw new Error("aborted after reading the zip index");
    }
    if (all.length > limits.maxEntries) {
      throw new FormatIQError("FILE_TOO_LARGE", "The file has too many parts.");
    }
    const entries = new Map<string, FileEntry>();
    let totalBytes = 0;
    for (const entry of all) {
      if (!isSafeZipPath(entry.filename)) {
        throw new FormatIQError("UNSAFE_ZIP_PATH", "The file has an unsafe path.", entry.filename);
      }
      if (entry.directory) continue;
      if (entries.has(entry.filename)) {
        throw new FormatIQError("ZIP_READ_FAILED", "The file has a duplicate part.", entry.filename);
      }
      if (entry.encrypted) {
        throw new FormatIQError("ZIP_READ_FAILED", "Password-protected files are not supported.");
      }
      totalBytes += entry.uncompressedSize;
      if (totalBytes > limits.maxTotalBytes) {
        throw new FormatIQError("FILE_TOO_LARGE", "The file is too big once unzipped.");
      }
      entries.set(entry.filename, entry);
    }
    const mainPart = entries.get("word/document.xml");
    if (!entries.has("[Content_Types].xml") || mainPart === undefined) {
      throw new FormatIQError("NOT_DOCX", "This does not look like a Word document.");
    }
    if (mainPart.uncompressedSize > limits.maxXmlBytes) {
      throw new FormatIQError("XML_TOO_LARGE", "The main document text is too large.");
    }
    return { entries, totalBytes, close: () => reader.close() };
  } catch (err) {
    await reader.close().catch(() => undefined);
    throw mapError(err, control, "ZIP_READ_FAILED");
  }
}

/**
 * Stream one entry in chunks. Stops with FILE_TOO_LARGE if more than
 * maxBytes come out (the zip index can lie, so we count real bytes).
 * Tested with zip.js 2.23.0: getData accepts a WritableStream.
 */
export async function streamEntry(
  entry: FileEntry,
  onChunk: (chunk: Uint8Array) => void,
  maxBytes: number,
  control?: JobControl,
): Promise<void> {
  let seen = 0;
  let limitError: FormatIQError | undefined;
  const sink = new WritableStream<Uint8Array>({
    write(chunk): void {
      seen += chunk.byteLength;
      if (seen > maxBytes) {
        limitError = new FormatIQError("FILE_TOO_LARGE", "A part is larger than allowed.", entry.filename);
        throw limitError;
      }
      onChunk(chunk);
    },
  });
  try {
    await entry.getData(sink, signalOptions(control));
  } catch (err) {
    throw limitError ?? mapError(err, control, "ZIP_READ_FAILED");
  }
}

/** Read one entry fully as UTF-8 text, with a size cap. */
export async function readEntryText(
  entry: FileEntry,
  maxBytes: number,
  control?: JobControl,
): Promise<string> {
  const decoder = new TextDecoder("utf-8");
  const pieces: string[] = [];
  await streamEntry(
    entry,
    (chunk) => {
      pieces.push(decoder.decode(chunk, { stream: true }));
    },
    maxBytes,
    control,
  );
  pieces.push(decoder.decode());
  return pieces.join("");
}

// ---------------------------------------------------------------------------
// 0d: writing
// ---------------------------------------------------------------------------

/** A zip being written. Call finish() last, once. */
export interface ZipOutput {
  /** Add a new or changed text part (for example edited document.xml). */
  addText(path: string, text: string, control?: JobControl): Promise<void>;
  /**
   * Add a part whose bytes are in a Blob, for example a big rewritten
   * document.xml collected as pieces: new Blob(pieces). Compressed like text
   * unless the name says it is already compressed (images and so on).
   */
  addBlob(path: string, blob: Blob, control?: JobControl): Promise<void>;
  /** Copy a part from an opened docx without changing its content. */
  copyEntry(entry: FileEntry, control?: JobControl): Promise<void>;
  /** Write the zip index and close. */
  finish(): Promise<void>;
}

/** Files that are already compressed. We store them (level 0), not squeeze. */
function isAlreadyCompressed(path: string): boolean {
  return /\.(jpe?g|png|gif|webp|heic|mp3|mp4|mov|zip|docx|xlsx|pptx|pdf)$/i.test(path);
}

function checkOutputPath(path: string): void {
  if (!isSafeZipPath(path)) {
    throw new FormatIQError("UNSAFE_ZIP_PATH", "Refusing to write an unsafe path.", path);
  }
}

function makeOutput(zip: ZipWriter<unknown>, onProgress?: ProgressCallback): ZipOutput {
  const progressOptions = (name: string): { onprogress?: (done: number, total: number) => Promise<void> } =>
    onProgress === undefined
      ? {}
      : {
          onprogress: async (done: number, total: number): Promise<void> => {
            onProgress({
              phase: "writing",
              fraction: total > 0 ? Math.min(1, done / total) : 0,
              message: `Writing ${name}`,
            });
          },
        };

  return {
    async addText(path, text, control) {
      checkOutputPath(path);
      try {
        await zip.add(path, new TextReader(text), {
          level: 6,
          ...signalOptions(control),
          ...progressOptions(path),
        });
      } catch (err) {
        throw mapError(err, control, "ZIP_WRITE_FAILED");
      }
    },

    async addBlob(path, blob, control) {
      checkOutputPath(path);
      try {
        await zip.add(path, new BlobReader(blob), {
          level: isAlreadyCompressed(path) ? 0 : 6,
          ...signalOptions(control),
          ...progressOptions(path),
        });
      } catch (err) {
        throw mapError(err, control, "ZIP_WRITE_FAILED");
      }
    },

    async copyEntry(entry, control) {
      checkOutputPath(entry.filename);
      // One part at a time is held as a Blob (never the whole file).
      // Streaming read and write together deadlocked in tests with
      // zip.js 2.23 (pipe, TransformStream and a manual bridge all hung).
      try {
        const part = await entry.getData(new BlobWriter(), signalOptions(control));
        await zip.add(entry.filename, new BlobReader(part), {
          level: isAlreadyCompressed(entry.filename) ? 0 : 6,
          lastModDate: entry.lastModDate,
          ...signalOptions(control),
          ...progressOptions(entry.filename),
        });
      } catch (err) {
        throw mapError(err, control, "ZIP_WRITE_FAILED");
      }
    },

    async finish() {
      try {
        await zip.close();
      } catch (err) {
        throw mapError(err, undefined, "ZIP_WRITE_FAILED");
      }
    },
  };
}

/**
 * Write the zip into any WritableStream. For big files, pass a stream that
 * goes straight to disk (for example from showSaveFilePicker().createWritable()).
 * Tested with zip.js 2.23.0: ZipWriter accepts a WritableStream.
 */
export function createZipOutput(
  target: WritableStream<Uint8Array>,
  onProgress?: ProgressCallback,
): ZipOutput {
  return makeOutput(new ZipWriter<unknown>(target), onProgress);
}

/**
 * Write the zip in memory and get a Blob at the end. Fine for small and
 * medium files. A 1 GB file should use createZipOutput with a disk stream.
 */
export function createBlobZipOutput(onProgress?: ProgressCallback): {
  output: ZipOutput;
  getBlob: () => Promise<Blob>;
} {
  const blobWriter = new BlobWriter(DOCX_MIME);
  const output = makeOutput(new ZipWriter<Blob>(blobWriter), onProgress);
  return { output, getBlob: () => blobWriter.getData() };
}
