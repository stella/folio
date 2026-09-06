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
const TARGET_ATTRIBUTE = /\bTarget\s*=\s*(?<quote>["'])(?<value>[^"']*)\k<quote>/u;
const TARGET_MODE_ATTRIBUTE = /\bTargetMode\s*=\s*(?<quote>["'])(?<value>[^"']*)\k<quote>/u;
const CONTENT_TYPES_PATTERN = /^\[Content_Types\]\.xml$/iu;

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
  const repair: PackageReferenceRepair = { danglingRelationships: [], danglingOverrides: [] };

  for (const { path, xml } of await readRelationshipParts(zip)) {
    let dropped = false;
    const pruned = xml.replaceAll(RELATIONSHIP_ELEMENT, (element) => {
      const target = relationshipTarget(element, path);
      if (target === null || isPresent(present, target)) {
        return element;
      }
      repair.danglingRelationships.push(target);
      dropped = true;
      return "";
    });
    if (dropped) {
      writePart({ zip, path, xml: pruned, compressionLevel });
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
