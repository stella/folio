/**
 * Whether a file is a WordprocessingML package at all, decided the way the
 * Open Packaging Conventions say to find its main part: the package
 * relationship of the office-document type names the part, the part exists,
 * `[Content_Types].xml` gives it a WordprocessingML main-document type, and
 * its root is a `document` in a WordprocessingML namespace.
 *
 * The folio parser reads whatever `.docx` it is handed and reports a missing
 * main part as a warning, so an archive holding no document opened as an
 * empty one. Checked where a file enters the CLI, a ZIP of unrelated entries
 * is refused before a read reports it as an empty document or a save commits
 * it over a real one.
 *
 * Every name is resolved by namespace URI, never by prefix, and both the
 * Transitional and the Strict relationship and WordprocessingML namespaces are
 * accepted.
 */

import { Result } from "better-result";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import JSZip from "jszip";
import path from "node:path";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

const CONTENT_TYPES_PART = "[Content_Types].xml";
const PACKAGE_RELATIONSHIPS_PART = "_rels/.rels";
/** The one main part the folio parser reads. */
const READ_MAIN_DOCUMENT_PART = "word/document.xml";

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";

/** The office-document relationship type: Transitional, then Strict. */
const OFFICE_DOCUMENT_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
]);

const WORDPROCESSINGML_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

/** A document, a template, and their macro-enabled forms; Strict uses the same types. */
const MAIN_DOCUMENT_CONTENT_TYPES: ReadonlySet<string> = new Set([
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  "application/vnd.ms-word.document.macroEnabled.main+xml",
  "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
]);

/** `[Content_Types].xml` and `_rels/.rels` are indexes; this bounds their inflation. */
const MAX_PACKAGE_INDEX_BYTES = 1024 * 1024;
/** How much of the main part is inflated to find its root element. */
const MAX_ROOT_SCAN_BYTES = 64 * 1024;

/** Why a file is not a readable WordprocessingML package. */
export const INVALID_PACKAGE_REASONS = {
  notZip: "notZip",
  duplicatePartName: "duplicatePartName",
  contentTypesMissing: "contentTypesMissing",
  contentTypesInvalid: "contentTypesInvalid",
  relationshipsMissing: "relationshipsMissing",
  relationshipsInvalid: "relationshipsInvalid",
  mainRelationshipMissing: "mainRelationshipMissing",
  mainRelationshipAmbiguous: "mainRelationshipAmbiguous",
  mainPartMissing: "mainPartMissing",
  mainPartContentType: "mainPartContentType",
  mainPartRoot: "mainPartRoot",
  mainPartLocation: "mainPartLocation",
} as const;

export type InvalidPackageReason =
  (typeof INVALID_PACKAGE_REASONS)[keyof typeof INVALID_PACKAGE_REASONS];

type Refusal = { reason: InvalidPackageReason; message: string };

const refuse = (reason: InvalidPackageReason, message: string): Result<never, Refusal> =>
  Result.err({ reason, message });

type XmlElement = {
  namespace: string | undefined;
  localName: string;
  attributes: Readonly<Record<string, string>>;
  children: readonly XmlElement[];
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const orderedParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  parseAttributeValue: false,
  parseTagValue: false,
  processEntities: false,
  ignoreDeclaration: true,
  ignorePiTags: true,
});

const splitName = (qualified: string): { prefix: string; localName: string } => {
  const colon = qualified.indexOf(":");
  return colon === -1
    ? { prefix: "", localName: qualified }
    : { prefix: qualified.slice(0, colon), localName: qualified.slice(colon + 1) };
};

const stringAttributes = (value: unknown): Record<string, string> => {
  const attributes: Record<string, string> = {};
  if (!isRecord(value)) return attributes;
  for (const [name, attribute] of Object.entries(value)) {
    if (typeof attribute === "string") attributes[name] = attribute;
  }
  return attributes;
};

/** The element children of one ordered node list, with namespaces in scope resolved. */
const elementsOf = (nodes: unknown, scope: ReadonlyMap<string, string>): XmlElement[] => {
  if (!Array.isArray(nodes)) return [];
  const elements: XmlElement[] = [];
  for (const node of nodes) {
    if (!isRecord(node)) continue;
    const name = Object.keys(node).find((key) => key !== ":@" && key !== "#text");
    if (name === undefined) continue;
    const attributes = stringAttributes(node[":@"]);
    const inner = new Map(scope);
    for (const [attribute, value] of Object.entries(attributes)) {
      if (attribute === "xmlns") inner.set("", value);
      else if (attribute.startsWith("xmlns:")) inner.set(attribute.slice("xmlns:".length), value);
    }
    const { prefix, localName } = splitName(name);
    elements.push({
      namespace: inner.get(prefix),
      localName,
      attributes,
      children: elementsOf(node[name], inner),
    });
  }
  return elements;
};

/** The root element of a well-formed XML part, or `null`. */
const parseRoot = (xml: string): XmlElement | null => {
  if (/<!DOCTYPE/iu.test(xml) || XMLValidator.validate(xml) !== true) return null;
  const parsed: unknown = orderedParser.parse(xml);
  const [root] = elementsOf(parsed, new Map());
  return root ?? null;
};

const childrenNamed = (
  element: XmlElement,
  namespace: string,
  localName: string,
): readonly XmlElement[] =>
  element.children.filter(
    (child) => child.namespace === namespace && child.localName === localName,
  );

type PartIndex = ReadonlyMap<string, JSZip.JSZipObject>;

/** Part names compare case-insensitively (ECMA-376 Part 2, 9.1.1.1). */
const partKey = (name: string): string => name.toLowerCase();

/** Inflate at most `limit` bytes of an entry; `truncated` when it holds more. */
const inflatePrefix = (
  entry: JSZip.JSZipObject,
  limit: number,
): Promise<{ text: string; truncated: boolean }> =>
  new Promise((resolve, reject) => {
    const decoder = new TextDecoder();
    let text = "";
    let size = 0;
    const stream = entry.internalStream("uint8array");
    stream
      .on("data", (chunk: Uint8Array) => {
        const room = limit - size;
        size += chunk.length;
        text += decoder.decode(chunk.subarray(0, Math.max(0, room)), { stream: true });
        if (size > limit) {
          stream.pause();
          resolve({ text: text + decoder.decode(), truncated: true });
        }
      })
      .on("error", reject)
      .on("end", () => resolve({ text: text + decoder.decode(), truncated: false }))
      .resume();
  });

type ReadIndexOptions = {
  parts: PartIndex;
  name: string;
  missing: InvalidPackageReason;
  invalid: InvalidPackageReason;
};

const readIndexPart = async ({
  parts,
  name,
  missing,
  invalid,
}: ReadIndexOptions): Promise<Result<XmlElement, Refusal>> => {
  const entry = parts.get(partKey(name));
  if (entry === undefined) return refuse(missing, `it has no ${name}.`);
  const inflated = await Result.tryPromise({
    try: () => inflatePrefix(entry, MAX_PACKAGE_INDEX_BYTES),
    catch: () => ({ reason: invalid, message: `its ${name} cannot be inflated.` }),
  });
  if (inflated.isErr()) return Result.err(inflated.error);
  if (inflated.value.truncated) {
    return refuse(invalid, `its ${name} is over ${String(MAX_PACKAGE_INDEX_BYTES)} bytes.`);
  }
  const root = parseRoot(inflated.value.text);
  return root === null ? refuse(invalid, `its ${name} is not well-formed XML.`) : Result.ok(root);
};

/** The part a package-relative relationship target names, without its leading `/`. */
const resolvePackageTarget = (target: string): string | null => {
  const decoded = Result.try({
    try: () => decodeURIComponent(target),
    catch: () => null,
  });
  if (decoded.isErr()) return null;
  const resolved = path.posix.normalize(path.posix.join("/", decoded.value));
  return resolved === "/" ? null : resolved.slice(1);
};

const mainPartName = (relationships: XmlElement): Result<string, Refusal> => {
  if (
    relationships.namespace !== PACKAGE_RELATIONSHIPS_NAMESPACE ||
    relationships.localName !== "Relationships"
  ) {
    return refuse(
      INVALID_PACKAGE_REASONS.relationshipsInvalid,
      `its ${PACKAGE_RELATIONSHIPS_PART} is not a relationships part.`,
    );
  }
  const main = childrenNamed(relationships, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship").filter(
    ({ attributes }) =>
      OFFICE_DOCUMENT_RELATIONSHIP_TYPES.has(attributes["Type"] ?? "") &&
      attributes["TargetMode"] !== "External",
  );
  const [only, ...others] = main;
  if (only === undefined) {
    return refuse(
      INVALID_PACKAGE_REASONS.mainRelationshipMissing,
      "no package relationship names a main document part.",
    );
  }
  if (others.length > 0) {
    return refuse(
      INVALID_PACKAGE_REASONS.mainRelationshipAmbiguous,
      `${String(main.length)} package relationships name a main document part.`,
    );
  }
  const part = resolvePackageTarget(only.attributes["Target"] ?? "");
  return part === null
    ? refuse(
        INVALID_PACKAGE_REASONS.relationshipsInvalid,
        "the main document relationship has no usable target.",
      )
    : Result.ok(part);
};

type ContentTypeOptions = { contentTypes: XmlElement; part: string };

/** An `Override` for the part wins over the `Default` for its extension. */
const contentTypeOf = ({ contentTypes, part }: ContentTypeOptions): string | undefined => {
  const partName = partKey(`/${part}`);
  const override = childrenNamed(contentTypes, CONTENT_TYPES_NAMESPACE, "Override").find(
    ({ attributes }) => partKey(attributes["PartName"] ?? "") === partName,
  );
  if (override !== undefined) return override.attributes["ContentType"];
  const extension = partKey(path.posix.extname(part).slice(1));
  return childrenNamed(contentTypes, CONTENT_TYPES_NAMESPACE, "Default").find(
    ({ attributes }) => partKey(attributes["Extension"] ?? "") === extension,
  )?.attributes["ContentType"];
};

/** The first element's name and namespace declarations, read from its start tag. */
const rootOfPrefix = (xml: string): XmlElement | null => {
  const body = xml.replace(/^﻿?(?:\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->)*/u, "");
  const match = /^<([A-Za-z_][\w.:-]*)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*\/?>/u.exec(
    body,
  );
  if (match === null) return null;
  const [, name = "", attributeText = ""] = match;
  const scope = new Map<string, string>();
  for (const [, attribute = "", double, single] of attributeText.matchAll(
    /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/gu,
  )) {
    const value = double ?? single ?? "";
    if (attribute === "xmlns") scope.set("", value);
    else if (attribute.startsWith("xmlns:")) scope.set(attribute.slice("xmlns:".length), value);
  }
  const { prefix, localName } = splitName(name);
  return { namespace: scope.get(prefix), localName, attributes: {}, children: [] };
};

const ZIP_END_SIGNATURE = 0x06054b50;
const ZIP64_END_SIGNATURE = 0x06064b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const ZIP_CENTRAL_MEMBER_SIGNATURE = 0x02014b50;
const ZIP_END_LENGTH = 22;
const ZIP_CENTRAL_MEMBER_LENGTH = 46;
const ZIP_MAX_COMMENT_LENGTH = 0xffff;

/** Read central-directory names before JSZip stores members by name. */
const checkRawMemberNames = (bytes: Uint8Array): Result<void, Refusal> => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (
    let offset = bytes.length - ZIP_END_LENGTH;
    offset >= Math.max(0, bytes.length - ZIP_END_LENGTH - ZIP_MAX_COMMENT_LENGTH);
    offset -= 1
  ) {
    if (
      view.getUint32(offset, true) === ZIP_END_SIGNATURE &&
      offset + ZIP_END_LENGTH + view.getUint16(offset + 20, true) === bytes.length
    ) {
      end = offset;
      break;
    }
  }
  if (end < 0) {
    return refuse(INVALID_PACKAGE_REASONS.notZip, "it has no readable ZIP directory.");
  }

  let count = view.getUint16(end + 10, true);
  let size = view.getUint32(end + 12, true);
  let directoryEnd = end;
  if (count === 0xffff || size === 0xffffffff || view.getUint32(end + 16, true) === 0xffffffff) {
    const locator = end - 20;
    if (locator < 0 || view.getUint32(locator, true) !== ZIP64_LOCATOR_SIGNATURE) {
      return refuse(INVALID_PACKAGE_REASONS.notZip, "its ZIP64 directory is missing.");
    }
    const zip64Offset = Number(view.getBigUint64(locator + 8, true));
    if (
      !Number.isSafeInteger(zip64Offset) ||
      zip64Offset < 0 ||
      zip64Offset + 56 > locator ||
      view.getUint32(zip64Offset, true) !== ZIP64_END_SIGNATURE
    ) {
      return refuse(INVALID_PACKAGE_REASONS.notZip, "its ZIP64 directory is invalid.");
    }
    count = Number(view.getBigUint64(zip64Offset + 32, true));
    size = Number(view.getBigUint64(zip64Offset + 40, true));
    directoryEnd = zip64Offset;
  }
  if (!Number.isSafeInteger(count) || !Number.isSafeInteger(size) || size > directoryEnd) {
    return refuse(INVALID_PACKAGE_REASONS.notZip, "its ZIP directory is invalid.");
  }

  const seen = new Set<string>();
  let offset = directoryEnd - size;
  for (let index = 0; index < count; index += 1) {
    if (
      offset + ZIP_CENTRAL_MEMBER_LENGTH > directoryEnd ||
      view.getUint32(offset, true) !== ZIP_CENTRAL_MEMBER_SIGNATURE
    ) {
      return refuse(INVALID_PACKAGE_REASONS.notZip, "its ZIP directory is invalid.");
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const next = offset + ZIP_CENTRAL_MEMBER_LENGTH + nameLength + extraLength + commentLength;
    if (next > directoryEnd) {
      return refuse(INVALID_PACKAGE_REASONS.notZip, "its ZIP directory is invalid.");
    }
    const name = Buffer.from(
      bytes.subarray(
        offset + ZIP_CENTRAL_MEMBER_LENGTH,
        offset + ZIP_CENTRAL_MEMBER_LENGTH + nameLength,
      ),
    ).toString("hex");
    if (seen.has(name)) {
      return refuse(
        INVALID_PACKAGE_REASONS.duplicatePartName,
        "two ZIP entries have the same name.",
      );
    }
    seen.add(name);
    offset = next;
  }
  if (offset !== directoryEnd) {
    return refuse(INVALID_PACKAGE_REASONS.notZip, "its ZIP directory is invalid.");
  }
  return Result.ok();
};

const checkPackage = async (bytes: Uint8Array): Promise<Result<void, Refusal>> => {
  const rawNames = checkRawMemberNames(bytes);
  if (rawNames.isErr()) return rawNames;
  const zip = await Result.tryPromise({
    try: () => JSZip.loadAsync(bytes),
    catch: () => ({
      reason: INVALID_PACKAGE_REASONS.notZip,
      message: "it is not a readable ZIP archive.",
    }),
  });
  if (zip.isErr()) return Result.err(zip.error);
  const parts = new Map<string, JSZip.JSZipObject>();
  for (const [name, entry] of Object.entries(zip.value.files)) {
    if (entry.dir) continue;
    const key = partKey(name.replace(/^\//u, ""));
    if (parts.has(key)) {
      return refuse(
        INVALID_PACKAGE_REASONS.duplicatePartName,
        `two entries share the part name ${name}.`,
      );
    }
    parts.set(key, entry);
  }

  const contentTypes = await readIndexPart({
    parts,
    name: CONTENT_TYPES_PART,
    missing: INVALID_PACKAGE_REASONS.contentTypesMissing,
    invalid: INVALID_PACKAGE_REASONS.contentTypesInvalid,
  });
  if (contentTypes.isErr()) return Result.err(contentTypes.error);
  if (
    contentTypes.value.namespace !== CONTENT_TYPES_NAMESPACE ||
    contentTypes.value.localName !== "Types"
  ) {
    return refuse(
      INVALID_PACKAGE_REASONS.contentTypesInvalid,
      `its ${CONTENT_TYPES_PART} is not a content types part.`,
    );
  }
  const relationships = await readIndexPart({
    parts,
    name: PACKAGE_RELATIONSHIPS_PART,
    missing: INVALID_PACKAGE_REASONS.relationshipsMissing,
    invalid: INVALID_PACKAGE_REASONS.relationshipsInvalid,
  });
  if (relationships.isErr()) return Result.err(relationships.error);

  const part = mainPartName(relationships.value);
  if (part.isErr()) return Result.err(part.error);
  const entry = parts.get(partKey(part.value));
  if (entry === undefined) {
    return refuse(
      INVALID_PACKAGE_REASONS.mainPartMissing,
      `its main document part /${part.value} is missing.`,
    );
  }
  const contentType = contentTypeOf({ contentTypes: contentTypes.value, part: part.value });
  if (contentType === undefined || !MAIN_DOCUMENT_CONTENT_TYPES.has(contentType)) {
    return refuse(
      INVALID_PACKAGE_REASONS.mainPartContentType,
      `its main part /${part.value} has content type ${contentType ?? "(none)"}, not a WordprocessingML document.`,
    );
  }
  const prefix = await Result.tryPromise({
    try: () => inflatePrefix(entry, MAX_ROOT_SCAN_BYTES),
    catch: () => ({
      reason: INVALID_PACKAGE_REASONS.mainPartRoot,
      message: `its main part /${part.value} cannot be inflated.`,
    }),
  });
  if (prefix.isErr()) return Result.err(prefix.error);
  const root = rootOfPrefix(prefix.value.text);
  if (
    root === null ||
    root.localName !== "document" ||
    !WORDPROCESSINGML_NAMESPACES.has(root.namespace ?? "")
  ) {
    return refuse(
      INVALID_PACKAGE_REASONS.mainPartRoot,
      `its main part /${part.value} is not a WordprocessingML document.`,
    );
  }
  if (partKey(part.value) !== partKey(READ_MAIN_DOCUMENT_PART)) {
    return refuse(
      INVALID_PACKAGE_REASONS.mainPartLocation,
      `its main document is /${part.value}; folio reads it only at /${READ_MAIN_DOCUMENT_PART}.`,
    );
  }
  return Result.ok();
};

/**
 * Refuse, as `invalid_document`, a file that is not a WordprocessingML
 * package with a main document part folio reads.
 */
export const checkWordprocessingPackage = async (
  filePath: string,
  bytes: Uint8Array,
): Promise<Result<void, FolioCliError>> => {
  const checked = await checkPackage(bytes);
  if (checked.isOk()) return Result.ok();
  const { reason, message } = checked.error;
  return Result.err(
    cliError({
      code: FOLIO_CLI_ERROR_CODES.invalidDocument,
      message: `${filePath} is not a .docx package: ${message}`,
      hint: "Pass a .docx saved by a word processor or by folio.",
      details: { reason },
    }),
  );
};
