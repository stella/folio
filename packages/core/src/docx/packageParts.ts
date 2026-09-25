/**
 * What a repack carries out of the source package, and the guarantee that
 * nothing in the output points at a part the output does not hold.
 *
 * A `.docx` is an OPC package: `[Content_Types].xml` declares every part in it
 * and the `.rels` parts wire them together. Folio models a handful of those
 * parts (the document body, headers, footers, notes, numbering, styles,
 * comments) and rewrites them on save. Everything else — embedded objects,
 * media, custom XML, fonts, charts, macro projects, parts folio has never heard
 * of — belongs to the author, and a repack hands it back byte for byte.
 *
 * Folio never executes document content and never removes it. The document is
 * the user's; the only thing a save refuses is a package path that would let
 * the archive escape its own directory when a host writes it out, which is a
 * property of the package structure rather than of the content. A macro-enabled
 * document therefore stays macro-enabled, main-part content type included:
 * silently handing back a `.docm` stripped of its project is data loss dressed
 * up as safety.
 *
 * A refused path loses its relationships and its content-type override with it.
 * That is the invariant this module exists for: Word and LibreOffice both
 * refuse a package whose `document.xml.rels` points at nothing, so a repack
 * that removes a part without removing its references produces a file that
 * does not open.
 */

import type JSZip from "jszip";

import { resolveRelativePath } from "./relsParser";

const CONTENT_TYPE_ELEMENT = /<(?:Default|Override)\b[^>]*?\/?>/giu;
const PART_NAME_ATTRIBUTE = /\bPartName\s*=\s*(?<quote>["'])(?<value>[^"']*)\k<quote>/u;
const RELATIONSHIP_ELEMENT = /<Relationship\b[^>]*?(?:\/>|>\s*<\/Relationship>)/giu;
const RELATIONSHIP_ID_ATTRIBUTE = /\bId\s*=\s*(?<quote>["'])(?<value>[^"']*)\k<quote>/u;
const TARGET_ATTRIBUTE = /\bTarget\s*=\s*(?<quote>["'])(?<value>[^"']*)\k<quote>/u;
const TARGET_MODE_ATTRIBUTE = /\bTargetMode\s*=\s*(?<quote>["'])(?<value>[^"']*)\k<quote>/u;
const CONTENT_TYPES_PATTERN = /^\[Content_Types\]\.xml$/iu;

/** Escape a string for literal use inside a `RegExp`. */
const escapeRegExp = (value: string): string => value.replaceAll(/[.*+?^${}()|[\]\\]/gu, "\\$&");

/**
 * The XML part a `.rels` part describes: `word/_rels/header1.xml.rels` ->
 * `word/header1.xml`; the package-level `_rels/.rels` has no single owning
 * part (its relationships are resolved by type, never by `r:id`), so it maps
 * to `undefined`.
 */
const owningPartPath = (relsPath: string): string | undefined => {
  const slash = relsPath.lastIndexOf("/_rels/");
  if (slash === -1) {
    return undefined;
  }
  const directory = relsPath.slice(0, slash);
  const fileName = relsPath.slice(slash + "/_rels/".length).replace(/\.rels$/iu, "");
  return fileName ? `${directory ? `${directory}/` : ""}${fileName}` : undefined;
};

/** Normalize a package path for comparison: lower case, no leading slash. */
const normalizePartPath = (path: string): string =>
  (path.startsWith("/") ? path.slice(1) : path).toLowerCase();

/**
 * Decode `%XX` escapes in a part name. `decodeURI` throws on a malformed
 * escape, and a malformed target is exactly the input this has to survive, so
 * the replacement is done by hand.
 */
const decodePartName = (name: string): string =>
  name.replaceAll(/%[0-9A-Fa-f]{2}/gu, (escape) =>
    String.fromCharCode(Number.parseInt(escape.slice(1), 16)),
  );

/**
 * Whether an entry path would escape the package when a host writes the archive
 * to disk: absolute, backslash-separated, or climbing out through `..`. This is
 * the only thing a save refuses, and it is about the archive, not the document.
 *
 * The full repack removes such an entry along with its references; the
 * selective path bails to the full repack rather than overlay one part on top
 * of a package carrying one. Both read this predicate, so neither can drift
 * into passing what the other blocks.
 */
export const isUnsafePackagePath = (path: string): boolean => {
  if (!path || path.startsWith("/") || path.includes("\\")) {
    return true;
  }
  return path.split("/").some((segment) => segment === "..");
};

/** Drop the entries whose path a save refuses. Every other part survives. */
export const removeUnsafeEntries = (zip: JSZip): void => {
  for (const [path, file] of Object.entries(zip.files)) {
    if (!file.dir && isUnsafePackagePath(path)) {
      zip.remove(path);
    }
  }
};

const presentPartPaths = (zip: JSZip): Set<string> => {
  const present = new Set<string>();
  for (const [path, file] of Object.entries(zip.files)) {
    if (!file.dir) {
      present.add(normalizePartPath(path));
    }
  }
  return present;
};

const isPresent = (present: ReadonlySet<string>, partPath: string): boolean =>
  present.has(normalizePartPath(partPath)) ||
  present.has(normalizePartPath(decodePartName(partPath)));

/** The part a relationship names, or null when it points outside the package. */
const relationshipTarget = (element: string, relsPath: string): string | null => {
  if (TARGET_MODE_ATTRIBUTE.exec(element)?.groups?.["value"] === "External") {
    return null;
  }
  const target = TARGET_ATTRIBUTE.exec(element)?.groups?.["value"];
  return target === undefined ? null : resolveRelativePath(relsPath, target);
};

/** What reconciliation removed to make the package internally consistent. */
export type PackageReferenceRepair = {
  /** Parts a relationship named that the package does not hold. */
  danglingRelationships: string[];
  /** Parts a content-type override named that the package does not hold. */
  danglingOverrides: string[];
  /**
   * `<part>|<id>` pairs: an `r:*`-style attribute in `part` named a
   * relationship id that reconciliation just removed from `part`'s own
   * `.rels` (because that relationship's target was itself a dangling
   * reference). Dropping the relationship without also dropping the
   * attribute that names it would trade one dangling reference for another —
   * a `.rels` entry pointing nowhere becomes a content attribute pointing at
   * nothing in its own `.rels`, which is no more valid.
   */
  orphanedIdReferences: string[];
};

type RelationshipPart = {
  path: string;
  xml: string;
};

/**
 * Every `.rels` part with its text. Read one at a time rather than through
 * `Promise.all`: a package holds a handful of tiny relationship parts, and the
 * sequential read keeps the output order stable.
 */
const readRelationshipParts = async (zip: JSZip): Promise<RelationshipPart[]> => {
  const parts: RelationshipPart[] = [];
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !normalizePartPath(path).endsWith(".rels")) {
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- a handful of tiny parts, read in package order
    parts.push({ path, xml: await file.async("text") });
  }
  return parts;
};

type WritePartOptions = {
  zip: JSZip;
  path: string;
  xml: string;
  compressionLevel: number;
};

const writePart = ({ zip, path, xml, compressionLevel }: WritePartOptions): void => {
  zip.file(path, xml, {
    compression: "DEFLATE",
    compressionOptions: { level: compressionLevel },
  });
};

/**
 * Make the package internally consistent before it is written: every internal
 * relationship target and every content-type override must name a part the
 * output actually holds.
 *
 * A part is rewritten only when something was dropped from it, so a package
 * that needs no repair leaves byte for byte as it arrived. Running this on an
 * already-reconciled package therefore reports nothing and changes nothing,
 * which is what the repack tests assert about their own output.
 */
export const reconcilePackageReferences = async (
  zip: JSZip,
  compressionLevel: number,
): Promise<PackageReferenceRepair> => {
  const present = presentPartPaths(zip);
  const repair: PackageReferenceRepair = {
    danglingRelationships: [],
    danglingOverrides: [],
    orphanedIdReferences: [],
  };

  // Ids removed from each `.rels` part, keyed by the part that references
  // them via `r:id` (not by the `.rels` path) so the second pass below can
  // read each owning part exactly once even though nothing else in this
  // function needs that grouping.
  const removedIdsByOwningPart = new Map<string, Set<string>>();

  for (const { path, xml } of await readRelationshipParts(zip)) {
    let dropped = false;
    const pruned = xml.replaceAll(RELATIONSHIP_ELEMENT, (element) => {
      const target = relationshipTarget(element, path);
      if (target === null || isPresent(present, target)) {
        return element;
      }
      repair.danglingRelationships.push(target);
      dropped = true;
      const id = RELATIONSHIP_ID_ATTRIBUTE.exec(element)?.groups?.["value"];
      const owningPart = owningPartPath(path);
      if (id !== undefined && owningPart !== undefined) {
        const ids = removedIdsByOwningPart.get(owningPart) ?? new Set<string>();
        ids.add(id);
        removedIdsByOwningPart.set(owningPart, ids);
      }
      return "";
    });
    if (dropped) {
      writePart({ zip, path, xml: pruned, compressionLevel });
    }
  }

  // A part whose own `.rels` just lost an id it still names via `r:id` (or
  // `r:embed`, `r:link`, …) would otherwise carry a reference nothing
  // resolves — the same "package points at nothing" failure this module
  // exists to prevent, just moved from the `.rels` file into the part's own
  // content. Scrub only the removed ids' attributes; every other reference
  // in the part is untouched.
  for (const [partPath, ids] of removedIdsByOwningPart) {
    const file = zip.file(partPath);
    if (!file) {
      continue;
    }
    const idPattern = [...ids].map((id) => escapeRegExp(id)).join("|");
    const referencePattern = new RegExp(
      `\\s+r:[A-Za-z]+\\s*=\\s*(?<quote>["'])(?<id>${idPattern})\\k<quote>`,
      "gu",
    );
    const xml = await file.async("text");
    const removedIds = new Set<string>();
    const scrubbed = xml.replaceAll(referencePattern, (...args) => {
      const groups = args.at(-1) as { id?: string } | undefined;
      if (groups?.id !== undefined) {
        removedIds.add(groups.id);
      }
      return "";
    });
    if (removedIds.size > 0) {
      writePart({ zip, path: partPath, xml: scrubbed, compressionLevel });
      for (const id of removedIds) {
        repair.orphanedIdReferences.push(`${partPath}|${id}`);
      }
    }
  }

  const contentTypes = zip.file(CONTENT_TYPES_PATTERN)[0];
  if (!contentTypes) {
    return repair;
  }
  let droppedOverride = false;
  const prunedContentTypes = (await contentTypes.async("text")).replaceAll(
    CONTENT_TYPE_ELEMENT,
    (element) => {
      const partName = PART_NAME_ATTRIBUTE.exec(element)?.groups?.["value"];
      if (partName === undefined || isPresent(present, partName)) {
        return element;
      }
      repair.danglingOverrides.push(normalizePartPath(decodePartName(partName)));
      droppedOverride = true;
      return "";
    },
  );
  if (droppedOverride) {
    writePart({ zip, path: contentTypes.name, xml: prunedContentTypes, compressionLevel });
  }
  return repair;
};

/** Every `r:*`-style attribute value found anywhere in an XML part's text. */
const REFERENCE_ATTRIBUTE = /\br:[A-Za-z]+\s*=\s*(?<quote>["'])(?<id>[^"']*)\k<quote>/gu;

/**
 * What {@link checkPackageIntegrity} found wrong with a package. Both fields
 * are empty for a package where every internal relationship resolves and
 * every reference to one does too — the state {@link reconcilePackageReferences}
 * is meant to leave behind, and the state a save must never regress from.
 */
export type PackageIntegrityReport = {
  /**
   * `<part>|<id>` pairs: an `r:*`-style attribute in `part` names a
   * relationship id absent from `part`'s own `.rels` — the id was never
   * declared, or a repair (this module's or any other) dropped it without
   * also dropping the attribute that names it.
   */
  unresolvedReferences: string[];
  /**
   * `<relsPart>|<target>` pairs: an internal (non-external) relationship in
   * `relsPart` names a target the package does not hold.
   */
  danglingRelationshipTargets: string[];
};

/**
 * Verify the two invariants an OPC package must hold for a conforming
 * consumer to accept it: every `r:*` reference in an XML part resolves in
 * that part's own `.rels`, and every internal relationship target names a
 * part the package actually holds.
 *
 * A read-only counterpart to {@link reconcilePackageReferences} — that
 * function repairs a package by removing what does not resolve; this one
 * only reports what still does not, so a test can assert a save produced
 * nothing for it to find. Cheap enough to run after every repack in a test
 * or a corpus-gate invariant: one pass over the `.rels` parts building an id
 * census, one pass over the XML parts checking references against it.
 */
export const checkPackageIntegrity = async (zip: JSZip): Promise<PackageIntegrityReport> => {
  const present = presentPartPaths(zip);
  const report: PackageIntegrityReport = {
    unresolvedReferences: [],
    danglingRelationshipTargets: [],
  };

  const idsByOwningPart = new Map<string, Set<string>>();
  for (const { path, xml } of await readRelationshipParts(zip)) {
    const owningPart = owningPartPath(path);
    const ids = new Set<string>();
    for (const element of xml.match(RELATIONSHIP_ELEMENT) ?? []) {
      const id = RELATIONSHIP_ID_ATTRIBUTE.exec(element)?.groups?.["value"];
      if (id !== undefined) {
        ids.add(id);
      }
      const target = relationshipTarget(element, path);
      if (target !== null && !isPresent(present, target)) {
        report.danglingRelationshipTargets.push(`${path}|${target}`);
      }
    }
    if (owningPart !== undefined) {
      idsByOwningPart.set(owningPart, ids);
    }
  }

  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !normalizePartPath(path).endsWith(".xml")) {
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- a handful of XML parts, read in package order
    const xml = await file.async("text");
    const ids = idsByOwningPart.get(path);
    for (const match of xml.matchAll(REFERENCE_ATTRIBUTE)) {
      const id = match.groups?.["id"];
      if (id !== undefined && !ids?.has(id)) {
        report.unresolvedReferences.push(`${path}|${id}`);
      }
    }
  }

  return report;
};
