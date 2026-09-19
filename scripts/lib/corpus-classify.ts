/**
 * Decide whether a corpus file is a WordprocessingML package at all.
 *
 * Public test corpora carry files that are `.docx` only by name: truncated
 * archives, encrypted packages, spreadsheet or theme-only packages, and XML
 * deliberately broken to exercise a parser's error path. Those are not evidence
 * about folio, so the gate classifies them out before any invariant runs.
 *
 * The check is deliberately independent of folio: it reads the container, not
 * the document, and it finds the main part the way OPC defines it — through the
 * package relationship, then by the root element's namespace and local name,
 * never by the conventional `word/document.xml` path. Word Online names that
 * part `word/document2.xml`, so a gate that looked for the conventional path
 * would file real packages as "not a docx" and hide whatever folio does with
 * them.
 *
 * Anything a reader can open this way is in scope, including the packages Word
 * accepts but a strict validator would reject: those are the inputs this gate
 * exists to find.
 */

import { XMLValidator } from "fast-xml-parser";
import JSZip from "jszip";

export const NOT_A_DOCX_REASONS = {
  encryptedPackage: "encrypted-package",
  oleCompoundFile: "ole-compound-file",
  notAZip: "not-a-zip",
  unreadableArchive: "unreadable-archive",
  notAnOpcPackage: "not-an-opc-package",
  notAWordprocessingPackage: "not-a-wordprocessing-package",
  malformedDocumentXml: "malformed-document-xml",
} as const;

export type NotADocxReason = (typeof NOT_A_DOCX_REASONS)[keyof typeof NOT_A_DOCX_REASONS];

export type CorpusClassification =
  | { kind: "docx"; documentPart: string }
  | { kind: "not-a-docx"; reason: NotADocxReason; detail: string };

const CONTENT_TYPES_PART = "[Content_Types].xml";
const PACKAGE_RELATIONSHIPS_PART = "_rels/.rels";
/** The Transitional and Strict spellings of the relationship that names the main part. */
const OFFICE_DOCUMENT_RELATIONSHIPS: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
]);
const WORDPROCESSINGML_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);
const DOCUMENT_ROOT_LOCAL_NAME = "document";

/** OLE2 compound-file signature: an encrypted OOXML package or a legacy binary document. */
const OLE_SIGNATURE = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
/** `EncryptedPackage`, the CFB stream name an ECMA-376 encrypted package carries, in UTF-16LE. */
const ENCRYPTED_PACKAGE_STREAM = Buffer.from("EncryptedPackage", "utf16le");

/** `Relationship` by local name: the part is valid with any prefix bound to the OPC namespace. */
const RELATIONSHIP_RE = /<(?:[\w.-]+:)?Relationship\b[^>]*>/gu;
const ATTRIBUTE_RE = /([\w:.-]+)\s*=\s*"([^"]*)"/gu;
/** The first element start tag, after any XML declaration, comment or doctype. */
const ROOT_ELEMENT_RE = /<([\w:.-]+)((?:\s[^<>]*)?)\/?>/u;
const PROLOG_RE = /^(?:\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>[]*(?:\[[\s\S]*?\])?\s*>)*/u;

const startsWith = (bytes: Uint8Array, prefix: Uint8Array): boolean =>
  bytes.length >= prefix.length && prefix.every((byte, index) => bytes[index] === byte);

const errorMessage = (cause: unknown): string =>
  cause instanceof Error ? cause.message : String(cause);

const attributesOf = (tag: string): Map<string, string> => {
  const attributes = new Map<string, string>();
  for (const [, name, value] of tag.matchAll(ATTRIBUTE_RE)) {
    if (name !== undefined && value !== undefined) {
      attributes.set(name, value);
    }
  }
  return attributes;
};

/** The `officeDocument` relationship target, as a package-rooted part name. */
const mainPartName = (relationshipsXml: string): string | null => {
  for (const [tag] of relationshipsXml.matchAll(RELATIONSHIP_RE)) {
    const attributes = attributesOf(tag);
    const type = attributes.get("Type");
    if (type === undefined || !OFFICE_DOCUMENT_RELATIONSHIPS.has(type)) {
      continue;
    }
    const target = attributes.get("Target");
    if (target === undefined) {
      continue;
    }
    return target.replace(/^\/+/u, "");
  }
  return null;
};

/** The root element's namespace URI and local name, resolved through its own prefix declarations. */
const rootElement = (xml: string): { namespace: string; localName: string } | null => {
  const body = xml.slice(PROLOG_RE.exec(xml)?.[0].length ?? 0);
  const match = ROOT_ELEMENT_RE.exec(body);
  if (match?.[1] === undefined) {
    return null;
  }
  const qualifiedName = match[1];
  const attributes = attributesOf(match[2] ?? "");
  const separator = qualifiedName.indexOf(":");
  const prefix = separator === -1 ? null : qualifiedName.slice(0, separator);
  const localName = separator === -1 ? qualifiedName : qualifiedName.slice(separator + 1);
  const namespace = attributes.get(prefix === null ? "xmlns" : `xmlns:${prefix}`) ?? "";
  return { namespace, localName };
};

export const classifyCorpusFile = async (bytes: Uint8Array): Promise<CorpusClassification> => {
  if (startsWith(bytes, OLE_SIGNATURE)) {
    const encrypted = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).includes(
      ENCRYPTED_PACKAGE_STREAM,
    );
    return encrypted
      ? {
          kind: "not-a-docx",
          reason: NOT_A_DOCX_REASONS.encryptedPackage,
          detail: "ECMA-376 encrypted package",
        }
      : {
          kind: "not-a-docx",
          reason: NOT_A_DOCX_REASONS.oleCompoundFile,
          detail: "OLE compound file, not an OPC package",
        };
  }

  if (bytes.length < 2 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    return {
      kind: "not-a-docx",
      reason: NOT_A_DOCX_REASONS.notAZip,
      detail: "missing the ZIP signature",
    };
  }

  let archive: JSZip;
  try {
    archive = await JSZip.loadAsync(bytes);
  } catch (cause) {
    return {
      kind: "not-a-docx",
      reason: NOT_A_DOCX_REASONS.unreadableArchive,
      detail: errorMessage(cause),
    };
  }

  const relationships = archive.file(PACKAGE_RELATIONSHIPS_PART);
  if (archive.file(CONTENT_TYPES_PART) === null || relationships === null) {
    return {
      kind: "not-a-docx",
      reason: NOT_A_DOCX_REASONS.notAnOpcPackage,
      detail: `no ${CONTENT_TYPES_PART} or ${PACKAGE_RELATIONSHIPS_PART}`,
    };
  }

  let mainPart: string | null;
  let documentXml: string;
  try {
    mainPart = mainPartName(await relationships.async("string"));
    if (mainPart === null) {
      return {
        kind: "not-a-docx",
        reason: NOT_A_DOCX_REASONS.notAnOpcPackage,
        detail: "no officeDocument relationship",
      };
    }
    const documentPart = archive.file(mainPart);
    if (documentPart === null) {
      return {
        kind: "not-a-docx",
        reason: NOT_A_DOCX_REASONS.notAWordprocessingPackage,
        detail: `the officeDocument relationship targets a missing part (${mainPart})`,
      };
    }
    documentXml = await documentPart.async("string");
  } catch (cause) {
    return {
      kind: "not-a-docx",
      reason: NOT_A_DOCX_REASONS.unreadableArchive,
      detail: errorMessage(cause),
    };
  }

  const validation = XMLValidator.validate(documentXml);
  if (validation !== true) {
    return {
      kind: "not-a-docx",
      reason: NOT_A_DOCX_REASONS.malformedDocumentXml,
      detail: validation.err.msg,
    };
  }

  const root = rootElement(documentXml);
  if (
    root === null ||
    root.localName !== DOCUMENT_ROOT_LOCAL_NAME ||
    !WORDPROCESSINGML_NAMESPACES.has(root.namespace)
  ) {
    return {
      kind: "not-a-docx",
      reason: NOT_A_DOCX_REASONS.notAWordprocessingPackage,
      detail: `${mainPart} is not a WordprocessingML document root`,
    };
  }

  return { kind: "docx", documentPart: mainPart };
};
