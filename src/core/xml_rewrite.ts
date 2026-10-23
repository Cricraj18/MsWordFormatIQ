/**
 * formatIQ raw XML rewriter (file 3a-1).
 *
 * xml.ts removes name prefixes ("w:val" -> "val"), so it cannot write XML
 * back out. This file keeps the full qualified names, so a part such as
 * word/document.xml can be read, changed token by token, and written again.
 *
 * Streaming: feed text with write(), call end() once. Output goes to the
 * onOutput callback in small pieces. Nothing large is kept in memory.
 *
 * Safety (same rules as xml.ts): any DOCTYPE throws XML_ENTITY_EXPANSION,
 * damaged XML throws ZIP_READ_FAILED. Only the five built-in entities and
 * numeric ones are decoded. Text and attributes are escaped again on output.
 *
 * Kept as they were: tag names, attribute order, namespace declarations,
 * comments, processing instructions (including the XML declaration) and
 * CDATA (CDATA comes out as ordinary escaped text, which means the same).
 * An empty element pair <a></a> comes out as <a/>. Same meaning in XML.
 */
import { parser as createSaxParser } from "sax";
import { FormatIQError } from "../types";

/** One attribute: qualified name and decoded value. Order is kept. */
export type RawAttr = readonly [name: string, value: string];

/** Opening tag. Self-closing tags give an "open" and then a "close" token. */
export interface RawOpen {
  kind: "open";
  name: string;
  attrs: RawAttr[];
}

/** Closing tag. */
export interface RawClose {
  kind: "close";
  name: string;
}

/** Text (already decoded). It is escaped again when written. */
export interface RawText {
  kind: "text";
  text: string;
}

/** Already-written XML copied as is (comment, processing instruction). */
export interface RawVerbatim {
  kind: "raw";
  xml: string;
}

export type RawToken = RawOpen | RawClose | RawText | RawVerbatim;

/** Receives each token from the parser and passes tokens on with emit(). */
export type RawTransform = (token: RawToken, emit: (token: RawToken) => void) => void;

/** Writes the text content of an element or attribute safely. */
export function escapeXmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Attribute values also escape quotes, tabs and line breaks. */
export function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\t/g, "&#9;")
    .replace(/\n/g, "&#10;")
    .replace(/\r/g, "&#13;");
}

/** Read one attribute value, or undefined if the tag has none by that name. */
export function getAttr(tag: RawOpen, name: string): string | undefined {
  for (const [key, value] of tag.attrs) {
    if (key === name) return value;
  }
  return undefined;
}

/** Copy of the tag with one attribute set (added at the end if new). */
export function withAttr(tag: RawOpen, name: string, value: string): RawOpen {
  let found = false;
  const attrs: RawAttr[] = tag.attrs.map((pair): RawAttr => {
    if (pair[0] === name) {
      found = true;
      return [name, value];
    }
    return pair;
  });
  if (!found) attrs.push([name, value]);
  return { kind: "open", name: tag.name, attrs };
}

/** Make an opening tag token. */
export function openTag(name: string, attrs: RawAttr[] = []): RawOpen {
  return { kind: "open", name, attrs };
}

/** Make a closing tag token. */
export function closeTag(name: string): RawClose {
  return { kind: "close", name };
}

function writeOpen(token: RawOpen, selfClose: boolean): string {
  let out = "<" + token.name;
  for (const [key, value] of token.attrs) {
    out += " " + key + '="' + escapeXmlAttr(value) + '"';
  }
  return out + (selfClose ? "/>" : ">");
}

/**
 * Turns tokens into text. An "open" token is held back one step: if a
 * matching "close" follows at once, <a/> is written instead of <a></a>.
 */
export class XmlSerializer {
  private pending: RawOpen | undefined;

  constructor(private readonly out: (chunk: string) => void) {}

  push(token: RawToken): void {
    if (this.pending !== undefined) {
      const open = this.pending;
      this.pending = undefined;
      if (token.kind === "close" && token.name === open.name) {
        this.out(writeOpen(open, true));
        return;
      }
      this.out(writeOpen(open, false));
    }
    switch (token.kind) {
      case "open":
        this.pending = token;
        break;
      case "close":
        this.out("</" + token.name + ">");
        break;
      case "text":
        if (token.text.length > 0) this.out(escapeXmlText(token.text));
        break;
      case "raw":
        this.out(token.xml);
        break;
    }
  }

  /** Writes anything still held back. Call once at the end. */
  flush(): void {
    if (this.pending !== undefined) {
      const open = this.pending;
      this.pending = undefined;
      this.out(writeOpen(open, false));
    }
  }
}

/** Feed text in pieces with write(), then call end() once. */
export interface XmlRewriter {
  write(text: string): void;
  end(): void;
}

/** Options for createXmlRewriter. */
export interface XmlRewriteOptions {
  /** Name of the part, used in error messages, for example "word/document.xml". */
  partName: string;
  /** Receives the rewritten XML in pieces. */
  onOutput: (chunk: string) => void;
  /** Changes tokens. Leave out to copy the XML unchanged. */
  transform?: RawTransform;
}

/**
 * Make a streaming rewriter.
 * Throws FormatIQError: XML_ENTITY_EXPANSION for any DOCTYPE, ZIP_READ_FAILED
 * for damaged XML. Errors from the transform function pass through unchanged.
 */
export function createXmlRewriter(options: XmlRewriteOptions): XmlRewriter {
  const sax = createSaxParser(true, { trim: false });
  const serializer = new XmlSerializer(options.onOutput);
  const emit = (token: RawToken): void => serializer.push(token);
  const handle: RawTransform = options.transform ?? ((token, out) => out(token));
  let failure: unknown;
  let inCdata = false;

  const fail = (error: unknown): void => {
    if (failure === undefined) failure = error;
  };
  const throwIfFailed = (): void => {
    if (failure !== undefined) throw failure;
  };
  const dispatch = (token: RawToken): void => {
    if (failure !== undefined) return;
    try {
      handle(token, emit);
    } catch (error) {
      fail(error);
    }
  };

  sax.onerror = (err: Error): void => {
    const firstLine = err.message.split("\n")[0] ?? "XML error";
    fail(
      new FormatIQError(
        "ZIP_READ_FAILED",
        "A part of the document has damaged XML.",
        `${options.partName}: ${firstLine}`,
      ),
    );
  };
  sax.ondoctype = (): void => {
    fail(new FormatIQError("XML_ENTITY_EXPANSION", "The file uses a DOCTYPE, which is not allowed.", options.partName));
  };
  sax.onopentag = (tag): void => {
    const attrs: RawAttr[] = [];
    for (const [key, value] of Object.entries(tag.attributes)) {
      attrs.push([key, typeof value === "string" ? value : value.value]);
    }
    dispatch({ kind: "open", name: tag.name, attrs });
  };
  sax.onclosetag = (name: string): void => {
    dispatch({ kind: "close", name });
  };
  sax.ontext = (text: string): void => {
    if (!inCdata) dispatch({ kind: "text", text });
  };
  sax.onopencdata = (): void => {
    inCdata = true;
  };
  sax.oncdata = (text: string): void => {
    dispatch({ kind: "text", text });
  };
  sax.onclosecdata = (): void => {
    inCdata = false;
  };
  sax.oncomment = (text: string): void => {
    dispatch({ kind: "raw", xml: "<!--" + text + "-->" });
  };
  sax.onprocessinginstruction = (pi: { name: string; body: string }): void => {
    const body = pi.body.length > 0 ? " " + pi.body : "";
    dispatch({ kind: "raw", xml: "<?" + pi.name + body + "?>" });
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
      serializer.flush();
    },
  };
}

/** Convenience for small parts and tests: whole string in, whole string out. */
export function rewriteXmlString(xml: string, partName: string, transform?: RawTransform): string {
  const pieces: string[] = [];
  const options: XmlRewriteOptions = { partName, onOutput: (chunk) => pieces.push(chunk) };
  if (transform !== undefined) options.transform = transform;
  const rewriter = createXmlRewriter(options);
  rewriter.write(xml);
  rewriter.end();
  return pieces.join("");
}
