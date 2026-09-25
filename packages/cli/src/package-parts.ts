/**
 * Package-level facts a transaction reports and checks: which parts a save
 * changed, the first revision id free for new revisions, and whether the
 * parts it changed are well formed with every relationship reference
 * resolved.
 */

import { Result } from "better-result";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import JSZip from "jszip";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

export type PartChange = "added" | "modified" | "removed";

export type ChangedPart = { part: string; change: PartChange };

const WORDPROCESSINGML_NAMESPACES = [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
];

const RELATIONSHIP_NAMESPACES = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

const loadZip = (bytes: Uint8Array): Promise<Result<JSZip, FolioCliError>> =>
  Result.tryPromise({
    try: () => JSZip.loadAsync(bytes),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.integrityFailed,
        message: `The package is not a readable ZIP archive: ${error instanceof Error ? error.message : String(error)}`,
      }),
  });

const partEntries = (zip: JSZip): Map<string, JSZip.JSZipObject> =>
  new Map(Object.entries(zip.files).filter(([, entry]) => !entry.dir));

const sameBytes = (left: Uint8Array, right: Uint8Array): boolean =>
  Buffer.compare(left, right) === 0;

/** Parts added, removed, or whose uncompressed content differs, sorted by name. */
export const diffPackages = async (
  before: Uint8Array,
  after: Uint8Array,
): Promise<Result<ChangedPart[], FolioCliError>> => {
  const beforeZip = await loadZip(before);
  if (beforeZip.isErr()) return Result.err(beforeZip.error);
  const afterZip = await loadZip(after);
  if (afterZip.isErr()) return Result.err(afterZip.error);
  const beforeParts = partEntries(beforeZip.value);
  const afterParts = partEntries(afterZip.value);
  const changes: ChangedPart[] = [];
  for (const [part, entry] of afterParts) {
    const previous = beforeParts.get(part);
    if (previous === undefined) {
      changes.push({ part, change: "added" });
      continue;
    }
    const [left, right] = await Promise.all([
      previous.async("uint8array"),
      entry.async("uint8array"),
    ]);
    if (!sameBytes(left, right)) {
      changes.push({ part, change: "modified" });
    }
  }
  for (const part of beforeParts.keys()) {
    if (!afterParts.has(part)) {
      changes.push({ part, change: "removed" });
    }
  }
  return Result.ok(changes.toSorted((left, right) => left.part.localeCompare(right.part)));
};

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");

/** Prefixes a part's namespace declarations bind to any of `namespaces`. */
const prefixesBoundTo = (xml: string, namespaces: readonly string[]): Set<string> => {
  const prefixes = new Set<string>();
  for (const match of xml.matchAll(/\sxmlns:([A-Za-z_][\w.-]*)\s*=\s*"([^"]*)"/gu)) {
    const [, prefix, uri] = match;
    if (prefix !== undefined && uri !== undefined && namespaces.includes(uri)) {
      prefixes.add(prefix);
    }
  }
  return prefixes;
};

/**
 * The first revision id above every WordprocessingML `id` attribute in the
 * package, so new revisions never collide with an existing one.
 */
export const nextRevisionIdSeed = async (
  bytes: Uint8Array,
): Promise<Result<number, FolioCliError>> => {
  const zip = await loadZip(bytes);
  if (zip.isErr()) return Result.err(zip.error);
  let highest = 0;
  for (const [part, entry] of partEntries(zip.value)) {
    if (!part.endsWith(".xml")) continue;
    const xml = await entry.async("string");
    for (const prefix of prefixesBoundTo(xml, WORDPROCESSINGML_NAMESPACES)) {
      const pattern = new RegExp(`\\s${escapeRegExp(prefix)}:id\\s*=\\s*"(\\d+)"`, "gu");
      for (const [, id] of xml.matchAll(pattern)) {
        highest = Math.max(highest, Number(id));
      }
    }
  }
  return Result.ok(highest + 1);
};

const relationshipParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  removeNSPrefix: true,
  isArray: (name) => name === "Relationship",
});

type Relationship = { id: string; target: string; external: boolean };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseRelationships = (xml: string): Relationship[] => {
  const parsed: unknown = relationshipParser.parse(xml);
  const root = isRecord(parsed) ? parsed["Relationships"] : undefined;
  const entries = isRecord(root) ? root["Relationship"] : undefined;
  if (!Array.isArray(entries)) return [];
  return entries.filter(isRecord).map((entry) => ({
    id: typeof entry["Id"] === "string" ? entry["Id"] : "",
    target: typeof entry["Target"] === "string" ? entry["Target"] : "",
    external: entry["TargetMode"] === "External",
  }));
};

/** `word/document.xml` -> `word/_rels/document.xml.rels`; `_rels/.rels` belongs to the package. */
const relationshipsPartFor = (part: string): string => {
  const directory = path.posix.dirname(part);
  const base = path.posix.basename(part);
  return directory === "." ? `_rels/${base}.rels` : `${directory}/_rels/${base}.rels`;
};

/** The part a `.rels` part describes (`""` for the package-level `_rels/.rels`). */
const sourcePartFor = (relsPart: string): string => {
  const directory = path.posix.dirname(path.posix.dirname(relsPart));
  const base = path.posix.basename(relsPart, ".rels");
  return directory === "." ? base : `${directory}/${base}`;
};

const resolveTarget = (relsPart: string, target: string): string => {
  if (target.startsWith("/")) return target.slice(1);
  const sourceDirectory = path.posix.dirname(sourcePartFor(relsPart));
  return path.posix.normalize(
    path.posix.join(sourceDirectory === "." ? "" : sourceDirectory, target),
  );
};

const integrityError = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.integrityFailed, message });

type CheckPartOptions = {
  zip: JSZip;
  parts: ReadonlyMap<string, JSZip.JSZipObject>;
  part: string;
};

const checkChangedPart = async ({ zip, parts, part }: CheckPartOptions): Promise<string | null> => {
  const entry = parts.get(part);
  if (entry === undefined || !(part.endsWith(".xml") || part.endsWith(".rels"))) {
    return null;
  }
  const xml = await entry.async("string");
  if (XMLValidator.validate(xml) !== true) {
    return `${part} is not well-formed XML.`;
  }
  if (part.endsWith(".rels")) {
    const lowerCaseParts = new Set([...parts.keys()].map((name) => name.toLowerCase()));
    for (const { id, target, external } of parseRelationships(xml)) {
      if (external) continue;
      const resolved = resolveTarget(part, target);
      if (!lowerCaseParts.has(resolved.toLowerCase())) {
        return `${part} relationship ${id} targets ${resolved}, which the package does not contain.`;
      }
    }
    return null;
  }
  const prefixes = prefixesBoundTo(xml, RELATIONSHIP_NAMESPACES);
  if (prefixes.size === 0) return null;
  const relsXml = await zip.file(relationshipsPartFor(part))?.async("string");
  const ids = new Set(parseRelationships(relsXml ?? "").map(({ id }) => id));
  for (const prefix of prefixes) {
    const pattern = new RegExp(`\\s${escapeRegExp(prefix)}:[A-Za-z]+\\s*=\\s*"([^"]*)"`, "gu");
    for (const [, id] of xml.matchAll(pattern)) {
      if (id !== undefined && !ids.has(id)) {
        return `${part} references relationship ${id}, which ${relationshipsPartFor(part)} does not define.`;
      }
    }
  }
  return null;
};

/**
 * Check what a save wrote before it replaces anything: the package opens,
 * every changed XML part is well formed, every relationship id a changed
 * part references resolves, every internal target a changed `.rels` part
 * names exists, and the package parses as a document again.
 */
export const checkPackageIntegrity = async (
  bytes: Uint8Array<ArrayBuffer>,
  changedParts: readonly ChangedPart[],
): Promise<Result<void, FolioCliError>> => {
  const zip = await loadZip(bytes);
  if (zip.isErr()) return Result.err(zip.error);
  const parts = partEntries(zip.value);
  for (const { part, change } of changedParts) {
    if (change === "removed") continue;
    const problem = await checkChangedPart({ zip: zip.value, parts, part });
    if (problem !== null) {
      return Result.err(integrityError(`The saved package failed validation: ${problem}`));
    }
  }
  const reopened = await Result.tryPromise({
    try: () => FolioDocxReviewer.fromBuffer(bytes.slice().buffer),
    catch: (error) =>
      integrityError(
        `The saved package does not reopen: ${error instanceof Error ? error.message : String(error)}`,
      ),
  });
  return reopened.isErr() ? Result.err(reopened.error) : Result.ok();
};
