/**
 * `folio sanitize`: strip what a document carries about its past before it
 * is reused as a template. Comments go with every part that describes them
 * (the comments part, `commentsExtended`, `commentsIds`,
 * `commentsExtensible`) and with their anchors and references in every
 * story; `people.xml` loses the authors no remaining revision names. Revision
 * save identifiers go from every part (`w:rsid*` attributes, `w:rsid` in
 * styles, and the `w:rsids` table in settings). Document properties naming
 * people or an organization go on request.
 */

import { panic, Result } from "better-result";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import {
  attributeNamespace,
  childElements,
  isElement,
  isW,
  localNameOf,
  mainDocumentPart,
  type OoxmlPackage,
  RELATIONSHIP_TYPES,
  removeElements,
  storyParts,
  walkElements,
  wAttribute,
  WORDPROCESSINGML,
  type XmlElement,
} from "./ooxml-package";

export const SANITIZE_TARGETS: readonly ["comments", "rsids", "properties"] = [
  "comments",
  "rsids",
  "properties",
] as const;

export type SanitizeTarget = (typeof SANITIZE_TARGETS)[number];

/** What `sanitize` removes when the caller names nothing. */
export const DEFAULT_SANITIZE_TARGETS: readonly SanitizeTarget[] = [
  "comments",
  "rsids",
];

const isSanitizeTarget = (value: unknown): value is SanitizeTarget =>
  SANITIZE_TARGETS.some((target) => target === value);

/** The comment parts a main document relates, besides people.xml. */
const COMMENT_PART_TYPES: ReadonlySet<string> = new Set([
  RELATIONSHIP_TYPES.comments,
  RELATIONSHIP_TYPES.commentsExtended,
  RELATIONSHIP_TYPES.commentsIds,
  RELATIONSHIP_TYPES.commentsExtensible,
]);

const COMMENT_MARKERS: ReadonlySet<string> = new Set([
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);

const W15_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.microsoft.com/office/word/2012/wordml",
]);

const DUBLIN_CORE: ReadonlySet<string> = new Set([
  "http://purl.org/dc/elements/1.1/",
]);
const CORE_PROPERTIES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
]);
const EXTENDED_PROPERTIES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  "http://purl.oclc.org/ooxml/officeDocument/extendedProperties",
]);

type PropertySpec = {
  namespaces: ReadonlySet<string>;
  localName: string;
  property: string;
};

/** Core properties naming a person. */
const CORE_PROPERTY_SPECS: readonly PropertySpec[] = [
  { namespaces: DUBLIN_CORE, localName: "creator", property: "creator" },
  {
    namespaces: CORE_PROPERTIES,
    localName: "lastModifiedBy",
    property: "lastModifiedBy",
  },
];

/** Extended properties naming an organization or a person. */
const EXTENDED_PROPERTY_SPECS: readonly PropertySpec[] = [
  {
    namespaces: EXTENDED_PROPERTIES,
    localName: "Company",
    property: "company",
  },
  {
    namespaces: EXTENDED_PROPERTIES,
    localName: "Manager",
    property: "manager",
  },
];

/** The properties removed from each document-properties part, by its relationship type. */
const PROPERTY_SPECS_BY_RELATIONSHIP: ReadonlyMap<
  string,
  readonly PropertySpec[]
> = new Map([
  [RELATIONSHIP_TYPES.coreProperties, CORE_PROPERTY_SPECS],
  [RELATIONSHIP_TYPES.strictCoreProperties, CORE_PROPERTY_SPECS],
  [RELATIONSHIP_TYPES.extendedProperties, EXTENDED_PROPERTY_SPECS],
]);

export type SanitizeCommentsReport = {
  /** Comments in the removed comments part. */
  comments: number;
  /** Range starts, range ends, and references removed from the stories. */
  anchors: number;
  /** Authors of the removed comments. */
  authors: string[];
  /** Entries removed from people.xml. */
  people: string[];
};

export type SanitizeRsidsReport = {
  /** `w:rsid*` attributes removed. */
  attributes: number;
  /** `w:rsids` tables and style `w:rsid` elements removed. */
  elements: number;
};

/** One removed document property, with the value it held. */
export type RemovedProperty = { part: string; property: string; value: string };

export type SanitizeReport = {
  remove: SanitizeTarget[];
  comments?: SanitizeCommentsReport;
  rsids?: SanitizeRsidsReport;
  properties?: RemovedProperty[];
};

const invalidInput = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message });

/** The targets an argument names, in canonical order; the default when it names none. */
const targetsFrom = (
  value: unknown,
): Result<readonly SanitizeTarget[], FolioCliError> => {
  if (value === undefined) return Result.ok(DEFAULT_SANITIZE_TARGETS);
  if (!Array.isArray(value) || value.length === 0) {
    return Result.err(
      invalidInput(
        `remove must list one or more of ${SANITIZE_TARGETS.join(", ")}.`,
      ),
    );
  }
  const unknown = value.filter((entry) => !isSanitizeTarget(entry));
  if (unknown.length > 0) {
    return Result.err(
      invalidInput(
        `remove names ${unknown.map((entry) => JSON.stringify(entry)).join(", ")}; expected ${SANITIZE_TARGETS.join(", ")}.`,
      ),
    );
  }
  return Result.ok(SANITIZE_TARGETS.filter((target) => value.includes(target)));
};

const sortedUnique = (values: Iterable<string>): string[] =>
  [...new Set(values)].toSorted();

/** Revision authors still named in a part: `w:author` on any element. */
const revisionAuthors = (root: XmlElement, into: Set<string>): void => {
  walkElements(root, (element) => {
    if (isW(element, "comment")) return;
    const author = wAttribute(element, "author");
    if (author !== undefined) into.add(author);
  });
};

/** Remove a comment reference, with its run when the run holds nothing else. */
const removeCommentMarkers = (root: XmlElement): number => {
  let anchors = 0;
  removeElements(root, (element) => {
    if (!isW(element, "r")) return false;
    const children = childElements(element);
    const references = children.filter((child) =>
      isW(child, "commentReference"),
    );
    if (references.length === 0) return false;
    const onlyReferences = children.every(
      (child) => isW(child, "commentReference") || isW(child, "rPr"),
    );
    if (!onlyReferences) return false;
    anchors += references.length;
    return true;
  });
  anchors += removeElements(root, (element) =>
    [...COMMENT_MARKERS].some((marker) => isW(element, marker)),
  );
  return anchors;
};

const sanitizeComments = (
  pkg: OoxmlPackage,
): Result<SanitizeCommentsReport, FolioCliError> => {
  const report: SanitizeCommentsReport = {
    comments: 0,
    anchors: 0,
    authors: [],
    people: [],
  };
  const main = mainDocumentPart(pkg);
  if (main.isErr()) return Result.err(main.error);
  if (main.value === null) return Result.ok(report);
  const relationships = pkg.relationships(main.value);
  if (relationships.isErr()) return Result.err(relationships.error);

  const authors = new Set<string>();
  const removedIds = new Set<string>();
  for (const { id, type, part } of relationships.value) {
    if (!COMMENT_PART_TYPES.has(type)) continue;
    removedIds.add(id);
    if (part === null) continue;
    if (type === RELATIONSHIP_TYPES.comments) {
      const comments = pkg.root(part);
      if (comments.isErr()) return Result.err(comments.error);
      for (const comment of childElements(comments.value ?? {})) {
        if (!isW(comment, "comment")) continue;
        report.comments += 1;
        const author = wAttribute(comment, "author");
        if (author !== undefined && author !== "") authors.add(author);
      }
    }
    const removed = pkg.remove(part);
    if (removed.isErr()) return Result.err(removed.error);
  }
  report.authors = sortedUnique(authors);

  const stories = storyParts(pkg);
  if (stories.isErr()) return Result.err(stories.error);
  const stillNamed = new Set<string>();
  for (const part of stories.value) {
    const story = pkg.root(part);
    if (story.isErr()) return Result.err(story.error);
    if (story.value === null) continue;
    const anchors = removeCommentMarkers(story.value);
    if (anchors > 0) pkg.touch(part);
    report.anchors += anchors;
    revisionAuthors(story.value, stillNamed);
  }

  const people = relationships.value.filter(
    ({ type }) => type === RELATIONSHIP_TYPES.people,
  );
  const removedPeople: string[] = [];
  for (const { id, part } of people) {
    if (part === null) continue;
    const peopleRoot = pkg.root(part);
    if (peopleRoot.isErr()) return Result.err(peopleRoot.error);
    if (peopleRoot.value === null) continue;
    const count = removeElements(peopleRoot.value, (element) => {
      if (!isElement(element, W15_NAMESPACES, "person")) return false;
      const author = Object.entries(element.attributes ?? {}).find(
        ([attribute]) =>
          localNameOf(attribute) === "author" &&
          W15_NAMESPACES.has(attributeNamespace(element, attribute) ?? ""),
      )?.[1];
      if (author === undefined || stillNamed.has(String(author))) return false;
      removedPeople.push(String(author));
      return true;
    });
    if (count === 0) continue;
    if (
      childElements(peopleRoot.value).some((element) =>
        isElement(element, W15_NAMESPACES, "person"),
      )
    ) {
      pkg.touch(part);
      continue;
    }
    removedIds.add(id);
    const removed = pkg.remove(part);
    if (removed.isErr()) return Result.err(removed.error);
  }
  report.people = sortedUnique(removedPeople);

  const unlinked = pkg.removeRelationships(main.value, removedIds);
  if (unlinked.isErr()) return Result.err(unlinked.error);
  return Result.ok(report);
};

const RSID_ATTRIBUTE = /^rsid/u;

const sanitizeRsids = (
  pkg: OoxmlPackage,
): Result<SanitizeRsidsReport, FolioCliError> => {
  const report: SanitizeRsidsReport = { attributes: 0, elements: 0 };
  for (const part of pkg.parts()) {
    // Only a part that spells the name can carry one; the rest stay unparsed.
    if (!(pkg.text(part)?.includes("rsid") ?? false)) continue;
    const root = pkg.root(part);
    if (root.isErr()) return Result.err(root.error);
    if (root.value === null) continue;
    let changed = 0;
    walkElements(root.value, (element) => {
      const attributes = Object.entries(element.attributes ?? {});
      const kept = attributes.filter(
        ([attribute]) =>
          !RSID_ATTRIBUTE.test(localNameOf(attribute)) ||
          !WORDPROCESSINGML.has(attributeNamespace(element, attribute) ?? ""),
      );
      const removed = attributes.length - kept.length;
      if (removed === 0) return;
      element.attributes = Object.fromEntries(kept);
      report.attributes += removed;
      changed += removed;
    });
    const elements = removeElements(
      root.value,
      (element) => isW(element, "rsids") || isW(element, "rsid"),
    );
    report.elements += elements;
    changed += elements;
    if (changed > 0) pkg.touch(part);
  }
  return Result.ok(report);
};

const removeProperties = (
  pkg: OoxmlPackage,
  part: string | null,
  specs: readonly PropertySpec[],
  removed: RemovedProperty[],
): Result<void, FolioCliError> => {
  if (part === null) return Result.ok(undefined);
  const root = pkg.root(part);
  if (root.isErr()) return Result.err(root.error);
  if (root.value === null) return Result.ok(undefined);
  const count = removeElements(root.value, (element) => {
    const spec = specs.find(({ namespaces, localName }) =>
      isElement(element, namespaces, localName),
    );
    if (spec === undefined) return false;
    const value = (element.elements ?? [])
      .map((node) => (node.type === "text" ? String(node.text ?? "") : ""))
      .join("");
    removed.push({ part, property: spec.property, value });
    return true;
  });
  if (count > 0) pkg.touch(part);
  return Result.ok(undefined);
};

const sanitizeProperties = (
  pkg: OoxmlPackage,
): Result<RemovedProperty[], FolioCliError> => {
  const relationships = pkg.relationships("");
  if (relationships.isErr()) return Result.err(relationships.error);
  const removed: RemovedProperty[] = [];
  for (const { type, part } of relationships.value) {
    const specs = PROPERTY_SPECS_BY_RELATIONSHIP.get(type);
    if (specs === undefined) continue;
    const done = removeProperties(pkg, part, specs, removed);
    if (done.isErr()) return Result.err(done.error);
  }
  return Result.ok(removed);
};

/** Remove what `args.remove` names (default: comments and rsids) and report it. */
export const sanitizePackage = (
  pkg: OoxmlPackage,
  args: Readonly<Record<string, unknown>>,
): Result<SanitizeReport, FolioCliError> => {
  const targets = targetsFrom(args["remove"]);
  if (targets.isErr()) return Result.err(targets.error);
  const report: SanitizeReport = { remove: [...targets.value] };
  for (const target of targets.value) {
    switch (target) {
      case "comments": {
        const comments = sanitizeComments(pkg);
        if (comments.isErr()) return Result.err(comments.error);
        report.comments = comments.value;
        break;
      }
      case "rsids": {
        const rsids = sanitizeRsids(pkg);
        if (rsids.isErr()) return Result.err(rsids.error);
        report.rsids = rsids.value;
        break;
      }
      case "properties": {
        const properties = sanitizeProperties(pkg);
        if (properties.isErr()) return Result.err(properties.error);
        report.properties = properties.value;
        break;
      }
      default: {
        const unreachable: never = target;
        return panic("Unhandled sanitize target", { unreachable });
      }
    }
  }
  return Result.ok(report);
};
