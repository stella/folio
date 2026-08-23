import type { Document } from "../types/document";
import { validateFolioDocumentModel } from "../docx/modelValidation";

export const AUTOSAVE_FORMAT_VERSION = 2;

type EncodedValue =
  | null
  | boolean
  | number
  | string
  | { type: "array"; value: EncodedValue[] }
  | { type: "object"; value: Record<string, EncodedValue> }
  | { type: "map"; value: [EncodedValue, EncodedValue][] }
  | { type: "date"; value: string }
  | { type: "arrayBuffer"; value: string }
  | { type: "undefined" };

export type DecodedAutoSave = {
  document: Document;
  savedAt: string;
  version: typeof AUTOSAVE_FORMAT_VERSION;
};

export type EncodedAutoSaveDocument = {
  json: string;
};

const BASE64_CHUNK_BYTES = 0x8000;

const encodeBase64 = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer);
  const chunks: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += BASE64_CHUNK_BYTES) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + BASE64_CHUNK_BYTES)));
  }
  return btoa(chunks.join(""));
};

const decodeBase64 = (value: string): ArrayBuffer => {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes.buffer;
};

const encodeValue = (value: unknown): EncodedValue => {
  if (value === undefined) {
    return { type: "undefined" };
  }
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  if (value instanceof Date) {
    return { type: "date", value: value.toISOString() };
  }
  if (value instanceof ArrayBuffer) {
    return { type: "arrayBuffer", value: encodeBase64(value) };
  }
  if (value instanceof Map) {
    return {
      type: "map",
      value: [...value.entries()].map(([key, entry]) => [encodeValue(key), encodeValue(entry)]),
    };
  }
  if (Array.isArray(value)) {
    return { type: "array", value: value.map(encodeValue) };
  }
  if (typeof value === "object") {
    const encoded: Record<string, EncodedValue> = {};
    for (const [key, entry] of Object.entries(value)) {
      encoded[key] = encodeValue(entry);
    }
    return { type: "object", value: encoded };
  }
  throw new TypeError(`Unsupported auto-save value: ${typeof value}`);
};

const decodeValue = (value: unknown): unknown => {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  if (typeof value !== "object" || !value || !("type" in value)) {
    throw new TypeError("Invalid auto-save encoded value");
  }

  const type = value.type;
  if (type === "undefined") {
    return undefined;
  }
  if (type === "date") {
    if (
      !("value" in value) ||
      typeof value.value !== "string" ||
      !Number.isFinite(Date.parse(value.value))
    ) {
      throw new TypeError("Invalid auto-save date");
    }
    return new Date(value.value);
  }
  if (type === "arrayBuffer") {
    if (!("value" in value) || typeof value.value !== "string") {
      throw new TypeError("Invalid auto-save binary data");
    }
    return decodeBase64(value.value);
  }
  if (type === "array") {
    if (!("value" in value) || !Array.isArray(value.value)) {
      throw new TypeError("Invalid auto-save array");
    }
    return value.value.map(decodeValue);
  }
  if (type === "map") {
    if (!("value" in value) || !Array.isArray(value.value)) {
      throw new TypeError("Invalid auto-save map");
    }
    return new Map(
      value.value.map((entry) => {
        if (!Array.isArray(entry) || entry.length !== 2) {
          throw new TypeError("Invalid auto-save map entry");
        }
        return [decodeValue(entry[0]), decodeValue(entry[1])];
      }),
    );
  }
  if (type === "object") {
    if (!("value" in value) || typeof value.value !== "object" || !value.value) {
      throw new TypeError("Invalid auto-save object");
    }
    const decoded: Record<string, unknown> = Object.create(null);
    for (const [key, entry] of Object.entries(value.value)) {
      if (key === "__proto__" || key === "constructor" || key === "prototype") {
        throw new TypeError("Invalid auto-save object key");
      }
      const decodedEntry = decodeValue(entry);
      if (decodedEntry !== undefined) {
        decoded[key] = decodedEntry;
      }
    }
    return decoded;
  }
  throw new TypeError("Unknown auto-save encoded value");
};

export const encodeAutoSaveDocument = (document: Document): EncodedAutoSaveDocument => {
  const snapshot = { ...document, originalBuffer: null };
  return { json: JSON.stringify(encodeValue(snapshot)) };
};

export const encodeAutoSaveEnvelopeFromDocument = (
  document: EncodedAutoSaveDocument,
  savedAt: string,
): string =>
  `{"document":${document.json},"savedAt":${JSON.stringify(savedAt)},"version":${AUTOSAVE_FORMAT_VERSION}}`;

export const encodeAutoSaveEnvelope = (document: Document, savedAt: string): string =>
  encodeAutoSaveEnvelopeFromDocument(encodeAutoSaveDocument(document), savedAt);

export const decodeAutoSaveEnvelope = (json: string): DecodedAutoSave | null => {
  try {
    const parsed: unknown = JSON.parse(json);
    if (
      typeof parsed !== "object" ||
      !parsed ||
      !("document" in parsed) ||
      !("savedAt" in parsed) ||
      !("version" in parsed)
    ) {
      return null;
    }
    if (
      parsed.version !== AUTOSAVE_FORMAT_VERSION ||
      typeof parsed.savedAt !== "string" ||
      !Number.isFinite(Date.parse(parsed.savedAt))
    ) {
      return null;
    }
    const document = decodeValue(parsed.document);
    if (typeof document !== "object" || !document) {
      return null;
    }
    // SAFETY: the model validator is the runtime boundary for this decoded value.
    const validation = validateFolioDocumentModel(document as Document);
    if (!validation.valid) {
      return null;
    }
    // SAFETY: validation above established the decoded value as a Document model.
    return {
      document: document as Document,
      savedAt: parsed.savedAt,
      version: AUTOSAVE_FORMAT_VERSION,
    };
  } catch {
    return null;
  }
};
