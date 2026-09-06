/**
 * The PDF object model: values, serialization, the cross-reference table and
 * the trailer.
 *
 * Everything a backend writes into a PDF file passes through here, which is
 * what makes byte-for-byte determinism enforceable rather than aspirational:
 * numbers go through one formatter, dictionaries carry their entries as an
 * ordered list (never an object or a `Map` whose iteration order depends on
 * how it was built), and object numbers come from one allocator in call
 * order.
 */

import { panic } from "better-result";
import { deflateSync } from "node:zlib";

/**
 * PDF reals carry roughly five significant digits; four decimals on a
 * coordinate in points is a third of a micron, far below anything a
 * rasterizer resolves. Fixing the precision here is what makes two runs that
 * computed the same coordinate by different arithmetic routes emit the same
 * bytes.
 */
const NUMBER_DECIMALS = 4;

/**
 * No page coordinate, font metric or matrix entry folio produces comes near
 * this. Past it `toFixed` switches to exponent notation, which PDF readers do
 * not accept, so a value this large is a producer bug and not something to
 * paint approximately.
 */
const MAX_ABS_NUMBER = 1e10;

/** The one number formatter. Every number in the file is emitted through it. */
export const formatNumber = (value: number): string => {
  if (!Number.isFinite(value)) {
    panic(`PDF numbers must be finite, got ${String(value)}`);
  }
  if (Math.abs(value) > MAX_ABS_NUMBER) {
    panic(`PDF number out of range: ${String(value)}`);
  }
  const trimmed = value.toFixed(NUMBER_DECIMALS).replace(/0+$/u, "").replace(/\.$/u, "");
  // A signed zero, or any value that rounds to zero from below, collapses to
  // "0": "-0" and "0" are the same point but different bytes.
  return trimmed === "" || trimmed === "-" || trimmed === "-0" ? "0" : trimmed;
};

export type PdfRef = { readonly kind: "ref"; readonly id: number };

export type PdfDictEntry = readonly [string, PdfValue];

export type PdfDict = { readonly kind: "dict"; readonly entries: readonly PdfDictEntry[] };

export type PdfStream = {
  readonly kind: "stream";
  readonly dict: PdfDict;
  readonly data: Uint8Array;
};

export type PdfValue =
  | { readonly kind: "null" }
  | { readonly kind: "bool"; readonly value: boolean }
  | { readonly kind: "number"; readonly value: number }
  | { readonly kind: "name"; readonly value: string }
  /** A byte string: every code unit must be < 256 (URIs, dates, CMap names). */
  | { readonly kind: "asciiString"; readonly value: string }
  /** A text string, emitted as UTF-16BE with a BOM so any script survives. */
  | { readonly kind: "textString"; readonly value: string }
  /** Pre-formatted hex digits, for values that are already bytes (`/ID`). */
  | { readonly kind: "hexString"; readonly value: string }
  | { readonly kind: "array"; readonly items: readonly PdfValue[] }
  | PdfDict
  | PdfStream
  | PdfRef;

export const pdfNull: PdfValue = { kind: "null" };
export const pdfBool = (value: boolean): PdfValue => ({ kind: "bool", value });
export const pdfNumber = (value: number): PdfValue => ({ kind: "number", value });
export const pdfName = (value: string): PdfValue => ({ kind: "name", value });
export const pdfAsciiString = (value: string): PdfValue => ({ kind: "asciiString", value });
export const pdfTextString = (value: string): PdfValue => ({ kind: "textString", value });
export const pdfHexString = (value: string): PdfValue => ({ kind: "hexString", value });
export const pdfArray = (items: readonly PdfValue[]): PdfValue => ({ kind: "array", items });
export const pdfNumberArray = (values: readonly number[]): PdfValue =>
  pdfArray(values.map(pdfNumber));

/**
 * Entries keep the order they are written in, and an `undefined` value drops
 * the key entirely: an optional entry is either present with a value or
 * absent, never `/Key null`.
 */
export const pdfDict = (
  entries: readonly (readonly [string, PdfValue | undefined])[],
): PdfDict => ({
  kind: "dict",
  entries: entries.filter((entry): entry is PdfDictEntry => entry[1] !== undefined),
});

export const pdfStream = (dict: PdfDict, data: Uint8Array): PdfStream => ({
  kind: "stream",
  dict,
  data,
});

/**
 * zlib's output depends on the compression level, so it is fixed rather than
 * left to the default: two runs of the writer must agree byte for byte, and
 * "whatever zlib felt like" is not an agreement.
 */
const DEFLATE_LEVEL = 9;

export const deflateForPdf = (data: Uint8Array): Uint8Array => {
  const deflated = deflateSync(data, { level: DEFLATE_LEVEL });
  return new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength);
};

/** A `/FlateDecode` stream: the one place raw bytes become a compressed one. */
export const pdfFlateStream = (
  entries: readonly (readonly [string, PdfValue | undefined])[],
  data: Uint8Array,
): PdfStream =>
  pdfStream(pdfDict([...entries, ["Filter", pdfName("FlateDecode")]]), deflateForPdf(data));

/** Regular characters may appear in a name unescaped; everything else is `#xx`. */
const NAME_REGULAR = /^[A-Za-z0-9._\-+]$/u;

const encodeName = (value: string): string => {
  let out = "/";
  for (const char of value) {
    out += NAME_REGULAR.test(char)
      ? char
      : `#${char.charCodeAt(0).toString(16).padStart(2, "0").toUpperCase()}`;
  }
  return out;
};

const encodeAsciiString = (value: string): string => {
  let out = "(";
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code > 0xff) {
      panic(`byte string carries a code point above 0xFF: ${value}`);
    }
    if (code === 0x28 || code === 0x29 || code === 0x5c) {
      out += `\\${String.fromCharCode(code)}`;
    } else if (code < 0x20 || code > 0x7e) {
      out += `\\${code.toString(8).padStart(3, "0")}`;
    } else {
      out += String.fromCharCode(code);
    }
  }
  return `${out})`;
};

const encodeTextString = (value: string): string => {
  let hex = "FEFF";
  for (let index = 0; index < value.length; index += 1) {
    hex += value.charCodeAt(index).toString(16).padStart(4, "0").toUpperCase();
  }
  return `<${hex}>`;
};

/** Collects the file's bytes and hands back the offset of each object. */
type ByteSink = {
  readonly writeAscii: (text: string) => void;
  readonly writeBytes: (bytes: Uint8Array) => void;
  readonly offset: () => number;
  readonly toBytes: () => Uint8Array;
};

const createByteSink = (): ByteSink => {
  const chunks: Uint8Array[] = [];
  let length = 0;
  const writeBytes = (bytes: Uint8Array) => {
    chunks.push(bytes);
    length += bytes.length;
  };
  return {
    writeAscii: (text) => {
      const bytes = new Uint8Array(text.length);
      for (let index = 0; index < text.length; index += 1) {
        bytes[index] = text.charCodeAt(index) & 0xff;
      }
      writeBytes(bytes);
    },
    writeBytes,
    offset: () => length,
    toBytes: () => {
      const out = new Uint8Array(length);
      let cursor = 0;
      for (const chunk of chunks) {
        out.set(chunk, cursor);
        cursor += chunk.length;
      }
      return out;
    },
  };
};

/** Serializes everything except a stream's payload, which is raw bytes. */
const serializeValue = (value: PdfValue): string => {
  switch (value.kind) {
    case "null":
      return "null";
    case "bool":
      return value.value ? "true" : "false";
    case "number":
      return formatNumber(value.value);
    case "name":
      return encodeName(value.value);
    case "asciiString":
      return encodeAsciiString(value.value);
    case "textString":
      return encodeTextString(value.value);
    case "hexString":
      return `<${value.value}>`;
    case "array":
      return `[${value.items.map(serializeValue).join(" ")}]`;
    case "dict":
      return `<<${value.entries.map(([key, entry]) => `${encodeName(key)} ${serializeValue(entry)}`).join(" ")}>>`;
    case "stream":
      // A stream is written by `writeObject`, which has the sink; reaching
      // here means one was nested inside an array or dictionary, which PDF
      // does not allow.
      return panic("a stream must be an indirect object, not a nested value");
    case "ref":
      return `${String(value.id)} 0 R`;
    default: {
      const unreachable: never = value;
      return panic(`unhandled PDF value: ${JSON.stringify(unreachable)}`);
    }
  }
};

export type PdfDocumentOptions = {
  readonly rootRef: PdfRef;
  readonly infoRef: PdfRef;
  /** Two copies of this hex digest become `/ID`. */
  readonly idHex: (bodyBytes: Uint8Array) => string;
};

export type PdfDocument = {
  /** Reserves an object number; the value may be supplied later. */
  readonly allocate: () => PdfRef;
  readonly define: (ref: PdfRef, value: PdfValue) => void;
  readonly add: (value: PdfValue) => PdfRef;
  readonly serialize: (options: PdfDocumentOptions) => Uint8Array;
};

const PDF_HEADER = "%PDF-1.7\n";

/**
 * A comment whose bytes are all above 127 tells any tool that transfers the
 * file that it is binary. The four bytes are the conventional ones.
 */
const BINARY_MARKER = new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]);

const XREF_FREE_ENTRY = "0000000000 65535 f \n";

export const createPdfDocument = (): PdfDocument => {
  const values = new Map<number, PdfValue>();
  let nextId = 1;

  const allocate = (): PdfRef => {
    const ref: PdfRef = { kind: "ref", id: nextId };
    nextId += 1;
    return ref;
  };

  const define = (ref: PdfRef, value: PdfValue) => {
    if (values.has(ref.id)) {
      panic(`PDF object ${String(ref.id)} defined twice`);
    }
    values.set(ref.id, value);
  };

  const writeObject = (sink: ByteSink, id: number, value: PdfValue) => {
    sink.writeAscii(`${String(id)} 0 obj\n`);
    if (value.kind === "stream") {
      const withLength = pdfDict([...value.dict.entries, ["Length", pdfNumber(value.data.length)]]);
      sink.writeAscii(`${serializeValue(withLength)}\nstream\n`);
      sink.writeBytes(value.data);
      sink.writeAscii("\nendstream");
    } else {
      sink.writeAscii(serializeValue(value));
    }
    sink.writeAscii("\nendobj\n");
  };

  return {
    allocate,
    define,
    add: (value) => {
      const ref = allocate();
      define(ref, value);
      return ref;
    },
    serialize: ({ rootRef, infoRef, idHex }) => {
      const sink = createByteSink();
      sink.writeAscii(PDF_HEADER);
      sink.writeBytes(BINARY_MARKER);

      const offsets = new Map<number, number>();
      for (let id = 1; id < nextId; id += 1) {
        const value = values.get(id);
        if (value === undefined) {
          panic(`PDF object ${String(id)} was allocated but never defined`);
        }
        offsets.set(id, sink.offset());
        writeObject(sink, id, value);
      }

      const xrefOffset = sink.offset();
      // The `/ID` hashes the body: it is a content fingerprint, so a second
      // run over the same display list and timestamp reproduces it exactly.
      const digest = idHex(sink.toBytes());

      sink.writeAscii(`xref\n0 ${String(nextId)}\n`);
      sink.writeAscii(XREF_FREE_ENTRY);
      for (let id = 1; id < nextId; id += 1) {
        const offset = offsets.get(id) ?? panic(`missing offset for PDF object ${String(id)}`);
        sink.writeAscii(`${String(offset).padStart(10, "0")} 00000 n \n`);
      }

      const trailer = pdfDict([
        ["Size", pdfNumber(nextId)],
        ["Root", rootRef],
        ["Info", infoRef],
        ["ID", pdfArray([pdfHexString(digest), pdfHexString(digest)])],
      ]);
      sink.writeAscii(
        `trailer\n${serializeValue(trailer)}\nstartxref\n${String(xrefOffset)}\n%%EOF\n`,
      );
      return sink.toBytes();
    },
  };
};
