/**
 * Package parts a paste copied in from another package: the parts kept markup
 * names (a chart, an embedding, a diagram's parts) and the relationships that
 * name them, in the main document's relationships or in those of the note,
 * comment, header or footer part holding the markup. The model carries such a part as bytes
 * until a save writes it.
 */

import { panic, Result } from "better-result";

import { escapeXmlAttribute } from "@stll/docx-core";
import { OP_STORIES, type OpStory } from "@stll/docx-core/ops";
import type { ZipArchive } from "@stll/docx-core/zip";

import type { Document, Relationship } from "../types/document";
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

/** The main document part, and its relationships. */
export const MAIN_DOCUMENT_PART = "word/document.xml";
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

/** A target as a part name: percent escapes decoded where they decode. */
export const decodedTarget = (target: string): string =>
  Result.try(() => decodeURIComponent(target)).unwrapOr(target);

/** The package path a relationship of the part at `partPath` targets. */
export const relationshipTargetPath = (
  partPath: string,
  target: string,
): string =>
  resolveRelativePath(relationshipsPartOf(partPath), decodedTarget(target));

/** A relationship, with the path of the part whose relationships hold it. */
export type OwnedRelationship = {
  relationship: Relationship | undefined;
  part: string | undefined;
};

/**
 * Whether two relationships name the same target in the same way: an
 * internal target resolved against the part each belongs to, since the same
 * relative target names another part from another folder.
 */
export const sameRelationship = (
  left: OwnedRelationship,
  right: OwnedRelationship,
): boolean => {
  const { relationship: one, part: onePart } = left;
  const { relationship: other, part: otherPart } = right;
  if (
    one === undefined ||
    other === undefined ||
    one.type !== other.type ||
    one.targetMode !== other.targetMode
  )
    return false;
  if (one.targetMode === "External") return one.target === other.target;
  return (
    onePart !== undefined &&
    otherPart !== undefined &&
    relationshipTargetPath(onePart, one.target).toLowerCase() ===
      relationshipTargetPath(otherPart, other.target).toLowerCase()
  );
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

/**
 * The model fields holding markup kept verbatim, whose office relationship
 * attributes name relationships of the part holding it: a drawing's or a
 * shape's capture, kept block and inline markup, an image's click and hover
 * links, and a shape's alternate content.
 */
export const KEPT_MARKUP_KEYS: ReadonlySet<string> = new Set([
  "rawXml",
  "xml",
  "verbatimXml",
  "hlinkHoverXml",
]);

/** Every piece of markup kept verbatim below `value`. */
const collectKeptMarkup = (value: unknown, found: string[]): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) collectKeptMarkup(item, found);
    return;
  }
  for (const [key, item] of Object.entries(value)) {
    if (
      KEPT_MARKUP_KEYS.has(key) &&
      typeof item === "string" &&
      item.startsWith("<")
    )
      found.push(item);
    else collectKeptMarkup(item, found);
  }
};

/** Every piece of markup kept verbatim in a model value: a note, a comment. */
export const keptMarkupIn = (value: unknown): string[] => {
  const found: string[] = [];
  collectKeptMarkup(value, found);
  return found;
};

type PublishCopiedPackagePartsOptions = {
  document: Document;
  zip: ZipArchive;
  compressionLevel: number;
};

/**
 * Write the parts a paste copied in, each under its content type, and the
 * relationships its kept markup names that the package does not hold yet:
 * the main document's for the body, and a note, comment, header or footer
 * part's own for the content in it. A part the package already holds stays
 * as it is.
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
  await appendNamedRelationships({
    zip,
    relsPath: MAIN_RELATIONSHIPS_PART,
    relationships: document.package.relationships,
    content: document.package.document.content,
    compressionLevel,
  });
  for (const { key, partPath, content } of storyParts(document)) {
    const relationships = document.package.annotationRelationships?.get(key);
    if (relationships === undefined || content.length === 0) continue;
    // The part the save writes the story into, and its relationships, in
    // the casing the package already uses for them.
    const relsPath = relationshipsPartOf(
      zip
        .paths()
        .find((path) => path.toLowerCase() === partPath.toLowerCase()) ??
        partPath,
    );
    // oxlint-disable-next-line no-await-in-loop -- each part owns its relationships
    await appendNamedRelationships({
      zip,
      relsPath:
        zip
          .paths()
          .find((path) => path.toLowerCase() === relsPath.toLowerCase()) ??
        relsPath,
      relationships,
      content,
      compressionLevel,
    });
  }
};

/**
 * Each part besides the main document whose relationships the model keeps,
 * with the content it holds: the notes, the comments, each header and footer.
 */
const storyParts = (document: Document) => {
  const parts: { key: string; partPath: string; content: unknown[] }[] = [
    {
      key: "footnotes",
      partPath: "word/footnotes.xml",
      content: document.package.footnotes ?? [],
    },
    {
      key: "endnotes",
      partPath: "word/endnotes.xml",
      content: document.package.endnotes ?? [],
    },
    {
      key: "comments",
      partPath: "word/comments.xml",
      content: document.package.document.comments ?? [],
    },
  ];
  for (const [kind, stories] of [
    ["header", document.package.headers],
    ["footer", document.package.footers],
  ] as const)
    for (const [rId, story] of stories ?? []) {
      const partPath = storyPartPath(document, { kind, rId });
      if (partPath !== undefined)
        parts.push({
          key: headerFooterRelationshipsKey(kind, rId),
          partPath,
          content: story.content,
        });
    }
  return parts;
};

/** The key a header's or footer's part relationships are kept under. */
export const headerFooterRelationshipsKey = (
  kind: "header" | "footer",
  rId: string,
): string => `${kind}:${rId}`;

/**
 * The relationships the ids of the story part kept under `key` resolve
 * through, as the parser reads them: the part's own, or, for a note or
 * comment part without its own, the main document's. A header or footer
 * part without its own has none: its ids restart in every part.
 */
export const partRelationshipsOf = (
  document: Document,
  key: string,
): ReadonlyMap<string, Relationship> | undefined =>
  document.package.annotationRelationships?.get(key) ??
  (isHeaderFooterRelationshipsKey(key)
    ? undefined
    : document.package.relationships);

/** Whether `key` names a header's or a footer's part relationships. */
export const isHeaderFooterRelationshipsKey = (key: string): boolean =>
  key.startsWith("header:") || key.startsWith("footer:");

/**
 * The key the model keeps a story part's own relationships under, in
 * `annotationRelationships`; undefined for the body, whose relationships are
 * the main document's.
 */
export const storyRelationshipsKey = (story: OpStory): string | undefined => {
  if (story === OP_STORIES.MAIN) return undefined;
  switch (story.kind) {
    case "footnote":
      return "footnotes";
    case "endnote":
      return "endnotes";
    case "header":
    case "footer":
      return headerFooterRelationshipsKey(story.kind, story.rId);
  }
};

/**
 * The path of the part holding `story`: a header's or a footer's as the main
 * document's relationships name it, undefined when they name none.
 */
export const storyPartPath = (
  document: Document,
  story: OpStory,
): string | undefined => {
  if (story === OP_STORIES.MAIN) return MAIN_DOCUMENT_PART;
  if (!("rId" in story)) return `word/${story.kind}s.xml`;
  const relationship = document.package.relationships?.get(story.rId);
  return relationship === undefined || relationship.targetMode === "External"
    ? undefined
    : mainRelationshipTarget(relationship.target);
};

type AppendNamedRelationshipsOptions = {
  zip: ZipArchive;
  relsPath: string;
  relationships: ReadonlyMap<string, Relationship> | undefined;
  /** The content whose kept markup names the relationships to write. */
  content: unknown;
  compressionLevel: number;
};

/**
 * Write into the relationships part at `relsPath` each relationship the
 * model holds and the part does not that `content`'s kept markup names:
 * markup copied in with its parts. A note or comment part with no
 * relationships yet receives a new relationships part.
 */
const appendNamedRelationships = async ({
  zip,
  relsPath,
  relationships,
  content,
  compressionLevel,
}: AppendNamedRelationshipsOptions): Promise<void> => {
  if (relationships === undefined || relationships.size === 0) return;
  const entry = zip.entry(relsPath);
  if (entry === undefined && relsPath === MAIN_RELATIONSHIPS_PART) return;
  const xml =
    entry === undefined
      ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE_URI}"></Relationships>`
      : decodeText(await entry.bytes(), { ignoreBOM: true });
  const held = parseRelationships(xml);
  const missing = [...relationships.values()].filter(({ id }) => !held.has(id));
  if (missing.length === 0) return;
  // A relationship the model holds and the part does not is written when the
  // content's kept markup names it: markup copied in with its parts.
  const markup = keptMarkupIn(content);
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
    parseXmlDocument(xml)?.name ?? panic("A relationships part has no root.");
  zip.set(
    relsPath,
    appendToXmlRoot(xml, { rootName, markup: additions.join("") }),
    { level: compressionLevel },
  );
};

/** The absolute package path a main document relationship targets. */
export const mainRelationshipTarget = (target: string): string =>
  resolveRelativePath(MAIN_RELATIONSHIPS_PART, target);
