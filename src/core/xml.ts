/**
 * formatIQ small XML reader (file 1a, part 1).
 *
 * Wraps the "sax" library: a streaming parser, so a 200 MB document.xml is
 * read piece by piece and never becomes one big object tree.
 *
 * Safety (tested): sax does NOT expand entities declared in a DOCTYPE, and we
 * also refuse any DOCTYPE at all (Word never writes one). Only the five
 * built-in entities (&amp; &lt; &gt; &quot; &apos;) and numeric ones work.
 *
 * Names: the prefix is removed ("w:val" becomes "val"). This is fine for
 * Word files, which use one prefix per namespace.
 *
 * Runs in the browser and in Node. Never touches the network.
 */
import { parser as createSaxParser } from "sax";
import type { FileEntry } from "@zip.js/zip.js";
import { FormatIQError } from "../types";
import { streamEntry } from "./zipio";
import type { JobControl } from "./zipio";

/** An opening tag, with attribute names and tag name stripped of prefixes. */
export interface XmlOpenTag {
  name: string;
  attrs: Readonly<Record<string, string>>;
}

/** Callbacks for the parser. All are optional. */
export interface XmlHandlers {
  /** A tag opened. Self-closing tags call open() and then close(). */
  open?(tag: XmlOpenTag): void;
  /** A tag closed. */
  close?(name: string): void;
  /** Text (or CDATA) inside a tag. May arrive in several pieces. */
  text?(text: string): void;
}

/** Feed text in pieces with write(), then call end() once. */
export interface XmlFeeder {
  write(text: string): void;
  end(): void;
}

/** "w:val" becomes "val". */
export function localName(qualifiedName: string): string {
  const colon = qualifiedName.indexOf(":");
  return colon === -1 ? qualifiedName : qualifiedName.slice(colon + 1);
}

/**
 * Make an incremental XML parser.
 * Throws FormatIQError: XML_ENTITY_EXPANSION for any DOCTYPE,
 * ZIP_READ_FAILED for damaged XML (message names the part).
 */
export function createXmlFeeder(handlers: XmlHandlers, partName: string): XmlFeeder {
  const sax = createSaxParser(true, { trim: false });
  let failure: FormatIQError | undefined;

  const fail = (error: FormatIQError): void => {
    if (failure === undefined) failure = error;
  };
  const throwIfFailed = (): void => {
    if (failure !== undefined) throw failure;
  };

  sax.onerror = (err: Error): void => {
    const firstLine = err.message.split("\n")[0] ?? "XML error";
    fail(new FormatIQError("ZIP_READ_FAILED", "A part of the document has damaged XML.", `${partName}: ${firstLine}`));
  };
  sax.ondoctype = (): void => {
    fail(new FormatIQError("XML_ENTITY_EXPANSION", "The file uses a DOCTYPE, which is not allowed.", partName));
  };
  sax.onopentag = (tag): void => {
    const attrs: Record<string, string> = {};
    for (const [key, value] of Object.entries(tag.attributes)) {
      attrs[localName(key)] = typeof value === "string" ? value : value.value;
    }
    handlers.open?.({ name: localName(tag.name), attrs });
  };
  sax.onclosetag = (name: string): void => {
    handlers.close?.(localName(name));
  };
  sax.ontext = (text: string): void => {
    handlers.text?.(text);
  };
  sax.oncdata = (text: string): void => {
    handlers.text?.(text);
  };

  return {
    write(text: string): void {
      throwIfFailed();
      sax.write(text);
      throwIfFailed();
    },
    end(): void {
      throwIfFailed();
      sax.close();
      throwIfFailed();
    },
  };
}

/**
 * Stream one zip part through the XML parser (UTF-8). Memory use stays small
 * because only one chunk is held at a time.
 * @param maxBytes Stops with FILE_TOO_LARGE if more bytes come out.
 */
export async function parseEntryXml(
  entry: FileEntry,
  handlers: XmlHandlers,
  maxBytes: number,
  control?: JobControl,
): Promise<void> {
  const feeder = createXmlFeeder(handlers, entry.filename);
  const decoder = new TextDecoder("utf-8");
  await streamEntry(
    entry,
    (chunk) => {
      feeder.write(decoder.decode(chunk, { stream: true }));
    },
    maxBytes,
    control,
  );
  const rest = decoder.decode();
  if (rest.length > 0) feeder.write(rest);
  feeder.end();
}
