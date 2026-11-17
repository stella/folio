/**
 * A WordprocessingML package held in memory for a command that rewrites its
 * XML parts directly: parts parse on first use, relationships resolve the
 * Open Packaging Conventions way, and only the parts a rewrite touched are
 * written back. Every element and attribute is matched by namespace URI plus
 * local name, never by prefix, in both the Transitional and the Strict
 * profile.
 */

import { Result } from "better-result";
import path from "node:path";

import {
  resolveNamespaceUri,
  XML_NAMESPACE_URI,
} from "@stll/docx-core/xml/namespaces";
import { ZipArchive } from "@stll/docx-core/zip";
import { resolvePackageRelationshipTarget } from "@stll/folio-core/docx/packageParts";
import {
  elementToXml,
  parseXml,
  type XmlElement,
} from "@stll/folio-core/docx/xmlParser";
import { decodeText } from "@stll/folio-core/utils/textEncoding";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

export type { XmlElement };

/** WordprocessingML main namespace: Transitional, then Strict. */
export const WORDPROCESSINGML: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

const PACKAGE_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const CONTENT_TYPES_PART = "[Content_Types].xml";
const PACKAGE_RELATIONSHIPS_PART = "_rels/.rels";

const TRANSITIONAL_RELATIONSHIP_PREFIX =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT_RELATIONSHIP_PREFIX =
  "http://purl.oclc.org/ooxml/officeDocument/relationships/";

type RelationshipName =
  | "officeDocument"
  | "header"
  | "footer"
  | "footnotes"
  | "endnotes"
  | "styles"
  | "numbering"
  | "settings"
  | "comments"
  | "extendedProperties"
  | "coreProperties"
  | "strictCoreProperties"
  | "commentsExtended"
  | "commentsIds"
  | "commentsExtensible"
  | "people";

/** Relationship types by their Transitional spelling; Strict ones are read as these. */
export const RELATIONSHIP_TYPES: Readonly<Record<RelationshipName, string>> = {
  officeDocument: `${TRANSITIONAL_RELATIONSHIP_PREFIX}officeDocument`,
  header: `${TRANSITIONAL_RELATIONSHIP_PREFIX}header`,
  footer: `${TRANSITIONAL_RELATIONSHIP_PREFIX}footer`,
  footnotes: `${TRANSITIONAL_RELATIONSHIP_PREFIX}footnotes`,
  endnotes: `${TRANSITIONAL_RELATIONSHIP_PREFIX}endnotes`,
  styles: `${TRANSITIONAL_RELATIONSHIP_PREFIX}styles`,
  numbering: `${TRANSITIONAL_RELATIONSHIP_PREFIX}numbering`,
  settings: `${TRANSITIONAL_RELATIONSHIP_PREFIX}settings`,
  comments: `${TRANSITIONAL_RELATIONSHIP_PREFIX}comments`,
  extendedProperties: `${TRANSITIONAL_RELATIONSHIP_PREFIX}extended-properties`,
  coreProperties:
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  strictCoreProperties:
    "http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/core-properties",
  commentsExtended:
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
  commentsIds:
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
  commentsExtensible:
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
  people: "http://schemas.microsoft.com/office/2011/relationships/people",
};

const normalizeRelationshipType = (type: string): string =>
  type.startsWith(STRICT_RELATIONSHIP_PREFIX)
    ? `${TRANSITIONAL_RELATIONSHIP_PREFIX}${type.slice(STRICT_RELATIONSHIP_PREFIX.length)}`
    : type;

/** One relationship of a part, with its internal target resolved to a part name. */
export type PackageRelationship = {
  id: string;
  /** The Transitional spelling of the type. */
  type: string;
  /** The part it names, or `null` for an external or unresolvable target. */
  part: string | null;
};

type ParsedPart = { prolog: string; host: XmlElement; root: XmlElement };

const XML_DECLARATION = /^\s*<\?xml[^?]*\?>/u;
const XML_PART = /\.(?:xml|rels)$/iu;

const malformed = (part: string, detail: string): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.invalidDocument,
    message: `The package part ${part} could not be read: ${detail}`,
  });

/** The local name of a qualified element or attribute name. */
export const localNameOf = (qualified: string): string => {
  const colon = qualified.indexOf(":");
  return colon === -1 ? qualified : qualified.slice(colon + 1);
};

const prefixOf = (qualified: string): string => {
  const colon = qualified.indexOf(":");
  return colon === -1 ? "" : qualified.slice(0, colon);
};

/** Whether `node` is an element named `localName` in one of `namespaces`. */
export const isElement = (
  node: XmlElement,
  namespaces: ReadonlySet<string>,
  localName: string,
): boolean =>
  node.type === "element" &&
  node.name !== undefined &&
  localNameOf(node.name) === localName &&
  node.namespaceUri !== undefined &&
  namespaces.has(node.namespaceUri);

/** Whether `node` is the WordprocessingML element `localName`. */
export const isW = (node: XmlElement, localName: string): boolean =>
  isElement(node, WORDPROCESSINGML, localName);

/** The namespace URI of an attribute of `element`; unprefixed attributes have none. */
export const attributeNamespace = (
  element: XmlElement,
  attribute: string,
): string | undefined => {
  const prefix = prefixOf(attribute);
  if (prefix === "" || prefix === "xmlns") return undefined;
  return prefix === "xml"
    ? XML_NAMESPACE_URI
    : resolveNamespaceUri(element.namespaceScope, prefix);
};

/** The source spelling of a WordprocessingML attribute of `element`, if present. */
export const wAttributeName = (
  element: XmlElement,
  localName: string,
): string | undefined =>
  Object.keys(element.attributes ?? {}).find(
    (attribute) =>
      localNameOf(attribute) === localName &&
      WORDPROCESSINGML.has(attributeNamespace(element, attribute) ?? ""),
  );

/** A WordprocessingML attribute's value, as a string. */
export const wAttribute = (
  element: XmlElement,
  localName: string,
): string | undefined => {
  const name = wAttributeName(element, localName);
  const value = name === undefined ? undefined : element.attributes?.[name];
  return value === undefined ? undefined : String(value);
};

/** The element children of `node`. */
export const childElements = (node: XmlElement): XmlElement[] =>
  (node.elements ?? []).filter((child) => child.type === "element");

/** Visit every element below `node` in document order, `node` included. */
export const walkElements = (
  node: XmlElement,
  visit: (element: XmlElement) => void,
): void => {
  if (node.type === "element") visit(node);
  for (const child of node.elements ?? []) walkElements(child, visit);
};

/**
 * Remove every element below `node` that `remove` selects, with its subtree,
 * and return how many were removed. The selected element's descendants are
 * not visited.
 */
export const removeElements = (
  node: XmlElement,
  remove: (element: XmlElement) => boolean,
): number => {
  if (node.elements === undefined) return 0;
  let removed = 0;
  const kept: XmlElement[] = [];
  for (const child of node.elements) {
    if (child.type === "element" && remove(child)) {
      removed += 1;
      continue;
    }
    removed += removeElements(child, remove);
    kept.push(child);
  }
  if (removed > 0) node.elements = kept;
  return removed;
};

const relationshipsPartFor = (part: string): string =>
  part === ""
    ? PACKAGE_RELATIONSHIPS_PART
    : path.posix.join(
        path.posix.dirname(part),
        "_rels",
        `${path.posix.basename(part)}.rels`,
      );

/** A package open for rewriting; see the module comment. */
export type OoxmlPackage = {
  /** Every part name, as the archive spells it. */
  readonly parts: () => readonly string[];
  /** The part named `part` (compared case-insensitively), as the archive spells it. */
  readonly find: (part: string) => string | undefined;
  /** The raw text of an XML part, or `undefined` when it is absent. */
  readonly text: (part: string) => string | undefined;
  /** The parsed root element of an XML part; `null` when it is absent. */
  readonly root: (part: string) => Result<XmlElement | null, FolioCliError>;
  /** Record that a parsed part's tree was changed, so it is written back. */
  readonly touch: (part: string) => void;
  /** Remove a part with its relationships part and its content-type override. */
  readonly remove: (part: string) => Result<void, FolioCliError>;
  /** The relationships of `part`; `""` for the package's own. */
  readonly relationships: (
    part: string,
  ) => Result<readonly PackageRelationship[], FolioCliError>;
  /** Remove relationships of `part` by id. */
  readonly removeRelationships: (
    part: string,
    ids: ReadonlySet<string>,
  ) => Result<void, FolioCliError>;
  /** The parts, sorted, a rewrite changed or removed so far. */
  readonly changedParts: () => readonly string[];
  /** The package with every changed part written back. */
  readonly toBytes: () => Promise<Uint8Array<ArrayBuffer>>;
};

/** Open a package for rewriting; every XML part's text is read up front. */
export const openOoxmlPackage = async (
  bytes: Uint8Array,
): Promise<Result<OoxmlPackage, FolioCliError>> => {
  const zip = await Result.tryPromise({
    try: () => ZipArchive.read(bytes),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `The package is not a readable ZIP archive: ${error instanceof Error ? error.message : String(error)}`,
      }),
  });
  if (zip.isErr()) return Result.err(zip.error);
  const archive = zip.value;
  const texts = new Map<string, string>();
  for (const entry of archive.entries()) {
    if (!XML_PART.test(entry.path)) continue;
    const read = await Result.tryPromise({
      try: () => entry.bytes(),
      catch: (error) =>
        malformed(
          entry.path,
          error instanceof Error ? error.message : String(error),
        ),
    });
    if (read.isErr()) return Result.err(read.error);
    texts.set(entry.path, decodeText(read.value));
  }
  const parsed = new Map<string, ParsedPart>();
  const touched = new Set<string>();
  const removed = new Set<string>();

  const find = (part: string): string | undefined => {
    const lower = part.toLowerCase();
    return archive
      .paths()
      .find((name) => name.toLowerCase() === lower && !removed.has(name));
  };

  const parsePart = (
    part: string,
  ): Result<ParsedPart | null, FolioCliError> => {
    const name = find(part);
    if (name === undefined) return Result.ok(null);
    const cached = parsed.get(name);
    if (cached !== undefined) return Result.ok(cached);
    const text = texts.get(name);
    if (text === undefined) return Result.err(malformed(name, "not XML"));
    const host = Result.try({
      try: () => parseXml(text),
      catch: (error) =>
        malformed(name, error instanceof Error ? error.message : String(error)),
    });
    if (host.isErr()) return Result.err(host.error);
    const root = childElements(host.value).at(0);
    if (root === undefined)
      return Result.err(malformed(name, "no root element"));
    const entry = {
      prolog: XML_DECLARATION.exec(text)?.at(0)?.trimStart() ?? "",
      host: host.value,
      root,
    };
    parsed.set(name, entry);
    return Result.ok(entry);
  };

  const root = (part: string): Result<XmlElement | null, FolioCliError> =>
    parsePart(part).map((entry) => entry?.root ?? null);

  const touch = (part: string): void => {
    const name = find(part);
    if (name !== undefined) touched.add(name);
  };

  const relationships = (
    part: string,
  ): Result<readonly PackageRelationship[], FolioCliError> => {
    const relsPart = relationshipsPartFor(part);
    const rels = root(relsPart);
    if (rels.isErr()) return Result.err(rels.error);
    if (rels.value === null) return Result.ok([]);
    const relsName = find(relsPart) ?? relsPart;
    const found: PackageRelationship[] = [];
    for (const element of childElements(rels.value)) {
      if (
        !isElement(
          element,
          new Set([PACKAGE_RELATIONSHIPS_NAMESPACE]),
          "Relationship",
        )
      )
        continue;
      const attributes = element.attributes ?? {};
      const id = attributes["Id"];
      const type = attributes["Type"];
      const target = attributes["Target"];
      if (typeof id !== "string" || typeof type !== "string") continue;
      const external = attributes["TargetMode"] === "External";
      const resolved =
        external || typeof target !== "string"
          ? undefined
          : resolvePackageRelationshipTarget(target, relsName);
      found.push({
        id,
        type: normalizeRelationshipType(type),
        part:
          resolved === undefined
            ? null
            : (find(resolved) ?? resolved.replace(/^\//u, "")),
      });
    }
    return Result.ok(found);
  };

  const removeRelationships = (
    part: string,
    ids: ReadonlySet<string>,
  ): Result<void, FolioCliError> => {
    if (ids.size === 0) return Result.ok(undefined);
    const relsPart = relationshipsPartFor(part);
    const rels = root(relsPart);
    if (rels.isErr()) return Result.err(rels.error);
    if (rels.value === null) return Result.ok(undefined);
    const count = removeElements(rels.value, (element) => {
      const id = element.attributes?.["Id"];
      return typeof id === "string" && ids.has(id);
    });
    if (count > 0) touch(relsPart);
    return Result.ok(undefined);
  };

  const removeContentTypeOverride = (
    part: string,
  ): Result<void, FolioCliError> => {
    const types = root(CONTENT_TYPES_PART);
    if (types.isErr()) return Result.err(types.error);
    if (types.value === null) return Result.ok(undefined);
    const partName = `/${part}`.toLowerCase();
    const count = removeElements(
      types.value,
      (element) =>
        isElement(element, new Set([CONTENT_TYPES_NAMESPACE]), "Override") &&
        String(element.attributes?.["PartName"] ?? "").toLowerCase() ===
          partName,
    );
    if (count > 0) touch(CONTENT_TYPES_PART);
    return Result.ok(undefined);
  };

  const remove = (part: string): Result<void, FolioCliError> => {
    const name = find(part);
    if (name === undefined) return Result.ok(undefined);
    const override = removeContentTypeOverride(name);
    if (override.isErr()) return override;
    const relsName = find(relationshipsPartFor(name));
    for (const gone of relsName === undefined ? [name] : [name, relsName]) {
      removed.add(gone);
      touched.delete(gone);
    }
    return Result.ok(undefined);
  };

  const changedParts = (): readonly string[] =>
    [...new Set([...touched, ...removed])].toSorted();

  const toBytes = async (): Promise<Uint8Array<ArrayBuffer>> => {
    for (const name of removed) archive.delete(name);
    for (const name of touched) {
      const entry = parsed.get(name);
      if (entry === undefined) continue;
      const body = (entry.host.elements ?? [])
        .filter((node) => node.type === "element" || node.type === "text")
        .map((node) => elementToXml(node))
        .join("");
      archive.set(name, `${entry.prolog}${body}`);
    }
    return await archive.generate();
  };

  return Result.ok({
    parts: () => archive.paths().filter((name) => !removed.has(name)),
    find,
    text: (part) => {
      const name = find(part);
      return name === undefined ? undefined : texts.get(name);
    },
    root,
    touch,
    remove,
    relationships,
    removeRelationships,
    changedParts,
    toBytes,
  });
};

/** The main document part the package relationship names, or `null`. */
export const mainDocumentPart = (
  pkg: OoxmlPackage,
): Result<string | null, FolioCliError> =>
  pkg
    .relationships("")
    .map(
      (relationships) =>
        relationships.find(
          ({ type }) => type === RELATIONSHIP_TYPES.officeDocument,
        )?.part ?? null,
    );

/**
 * The main document and every story part it relates: headers, footers,
 * footnotes, and endnotes, each once, in relationship order.
 */
export const storyParts = (
  pkg: OoxmlPackage,
): Result<readonly string[], FolioCliError> => {
  const main = mainDocumentPart(pkg);
  if (main.isErr()) return Result.err(main.error);
  if (main.value === null) return Result.ok([]);
  const relationships = pkg.relationships(main.value);
  if (relationships.isErr()) return Result.err(relationships.error);
  const storyTypes: ReadonlySet<string> = new Set([
    RELATIONSHIP_TYPES.header,
    RELATIONSHIP_TYPES.footer,
    RELATIONSHIP_TYPES.footnotes,
    RELATIONSHIP_TYPES.endnotes,
  ]);
  const parts = [main.value];
  for (const { type, part } of relationships.value) {
    if (part === null || !storyTypes.has(type) || parts.includes(part))
      continue;
    if (pkg.find(part) !== undefined) parts.push(part);
  }
  return Result.ok(parts);
};
