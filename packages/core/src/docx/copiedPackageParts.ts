/**
 * Package parts a paste copied in from another package: the parts kept markup
 * names (a chart, an embedding, a diagram's parts) and the relationships that
 * name them. The model carries such a part as bytes until a save writes it.
 */

import { panic, Result } from "better-result";

import { escapeXmlAttribute } from "@stll/docx-core";
import type { ZipArchive } from "@stll/docx-core/zip";

import type { Document } from "../types/document";
import { decodeText } from "../utils/textEncoding";
import { registerContentTypeParts } from "./contentTypeRegistry";
import { isUnsafePackagePath } from "./packageParts";
import { parseRelationships, resolveRelativePath } from "./relsParser";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findAttributeByNamespaceUri,
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  parseXmlDocument,
  type XmlElement,
} from "./xmlParser";
import { appendToXmlRoot } from "./xmlRootInsertion";

/** The main document part's relationships. */
export const MAIN_RELATIONSHIPS_PART = "word/_rels/document.xml.rels";

/** The content type of every relationships part. */
export const RELATIONSHIPS_CONTENT_TYPE =
  "application/vnd.openxmlformats-package.relationships+xml";

const PACKAGE_RELATIONSHIPS_NAMESPACE_URI =
  "http://schemas.openxmlformats.org/package/2006/relationships";

/** The relationships part of the part at `path`: `<dir>/_rels/<name>.rels`. */
export const relationshipsPartOf = (path: string): string => {
  const slash = path.lastIndexOf("/");
  return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
};

/**
 * A relationship target relative to the part at `partPath`, naming the part
 * at `absolute`: the inverse of {@link resolveRelativePath}.
 */
export const relativeTarget = (partPath: string, absolute: string): string => {
  const from = partPath.split("/").slice(0, -1);
  const to = absolute.split("/");
  let shared = 0;
  while (
    shared < from.length &&
    shared < to.length - 1 &&
    from[shared] === to[shared]
  )
    shared += 1;
  return `${"../".repeat(from.length - shared)}${to.slice(shared).join("/")}`;
};

/** Visit every office relationship attribute in `element` and below it. */
const visitRelationshipAttributes = (
  element: XmlElement,
  visit: (element: XmlElement, name: string, value: string) => void,
): void => {
  for (const name of Object.keys(element.attributes ?? {})) {
    const attribute = findAttributeByNamespaceUri(
      element,
      OFFICE_RELATIONSHIP_NAMESPACE_URIS,
      getLocalName(name),
    );
    if (attribute !== null && attribute.name === name)
      visit(element, name, attribute.value);
  }
  for (const child of element.elements ?? [])
    visitRelationshipAttributes(child, visit);
};

/** Kept markup read as the children of one wrapping element. */
const parsedMarkup = (xml: string) =>
  Result.try(() => {
    const wrapper = parseXml(`<root>${xml}</root>`, OOXML_NAMESPACE_SCOPE)
      .elements?.[0];
    if (wrapper === undefined) throw new TypeError("Kept markup is empty.");
    return wrapper;
  });

/**
 * The relationship ids kept markup names, in any office relationship
 * attribute; undefined when the markup cannot be read.
 */
export const relationshipIdsInMarkup = (xml: string): string[] | undefined => {
  const parsed = parsedMarkup(xml);
  if (parsed.isErr()) return undefined;
  const ids: string[] = [];
  visitRelationshipAttributes(parsed.value, (_element, _name, value) =>
    ids.push(value),
  );
  return ids;
};

/**
 * Kept markup with each relationship id `ids` maps renamed, as the same
 * markup otherwise; undefined when the markup cannot be read.
 */
export const withRelationshipIds = (
  xml: string,
  ids: ReadonlyMap<string, string>,
): string | undefined => {
  const parsed = parsedMarkup(xml);
  if (parsed.isErr()) return undefined;
  let renamed = false;
  visitRelationshipAttributes(parsed.value, (element, name, value) => {
    const next = ids.get(value);
    if (next === undefined || element.attributes === undefined) return;
    element.attributes[name] = next;
    renamed = true;
  });
  return renamed
    ? (parsed.value.elements ?? []).map(captureVerbatimXml).join("")
    : xml;
};

/** Every piece of markup kept verbatim below `value`. */
const collectKeptMarkup = (value: unknown, found: string[]): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) collectKeptMarkup(item, found);
    return;
  }
  for (const [key, item] of Object.entries(value)) {
    if (
      (key === "rawXml" || key === "xml") &&
      typeof item === "string" &&
      item.startsWith("<")
    )
      found.push(item);
    else collectKeptMarkup(item, found);
  }
};

type PublishCopiedPackagePartsOptions = {
  document: Document;
  zip: ZipArchive;
  compressionLevel: number;
};

/**
 * Write the parts a paste copied in, each under its content type, and the
 * main document relationships its kept markup names that the package does
 * not hold yet. A part the package already holds stays as it is.
 */
export const publishCopiedPackageParts = async ({
  document,
  zip,
  compressionLevel,
}: PublishCopiedPackagePartsOptions): Promise<void> => {
  const written: { partName: string; contentType: string }[] = [];
  const paths = new Set(zip.paths().map((path) => path.toLowerCase()));
  for (const part of document.package.copiedParts?.values() ?? []) {
    if (isUnsafePackagePath(part.path) || !part.path.startsWith("word/"))
      panic("A copied package part has an invalid package path.");
    if (paths.has(part.path.toLowerCase())) continue;
    zip.set(part.path, new Uint8Array(part.data), { level: compressionLevel });
    paths.add(part.path.toLowerCase());
    written.push({ partName: `/${part.path}`, contentType: part.contentType });
  }
  if (written.length > 0) {
    const registry =
      zip.entry("[Content_Types].xml") ??
      panic("A package receiving copied parts has no content types.");
    const xml = decodeText(await registry.bytes(), { ignoreBOM: true });
    zip.set(registry.path, registerContentTypeParts(xml, written), {
      level: compressionLevel,
    });
  }
  const relationships = document.package.relationships;
  const entry = zip.entry(MAIN_RELATIONSHIPS_PART);
  if (relationships === undefined || entry === undefined) return;
  const xml = decodeText(await entry.bytes(), { ignoreBOM: true });
  const held = parseRelationships(xml);
  const missing = [...relationships.values()].filter(({ id }) => !held.has(id));
  if (missing.length === 0) return;
  // A relationship the model holds and the part does not is written when the
  // body's kept markup names it: markup copied in with its parts.
  const markup: string[] = [];
  collectKeptMarkup(document.package.document.content, markup);
  const named = new Set(
    markup
      .filter((kept) => missing.some(({ id }) => kept.includes(id)))
      .flatMap((kept) => relationshipIdsInMarkup(kept) ?? []),
  );
  const additions = missing
    .filter(({ id }) => named.has(id))
    .map(
      ({ id, type, target, targetMode }) =>
        `<Relationship xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE_URI}" Id="${escapeXmlAttribute(id)}" Type="${escapeXmlAttribute(type)}" Target="${escapeXmlAttribute(target)}"${targetMode === "External" ? ' TargetMode="External"' : ""}/>`,
    );
  if (additions.length === 0) return;
  const rootName =
    parseXmlDocument(xml)?.name ??
    panic("The main document relationships have no root.");
  zip.set(
    entry.path,
    appendToXmlRoot(xml, { rootName, markup: additions.join("") }),
    { level: compressionLevel },
  );
};

/** The absolute package path a main document relationship targets. */
export const mainRelationshipTarget = (target: string): string =>
  resolveRelativePath(MAIN_RELATIONSHIPS_PART, target);
