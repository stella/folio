/**
 * What the stories of a package point at outside their own markup: the
 * relationships of the main document and of the note parts, the media those
 * name, and a fingerprint for an object that travels whole.
 */
import { Result } from "better-result";

import { resolveNamespaceUri } from "@stll/docx-core/xml/namespaces";

import type { RelationshipMap } from "../../../types/document";
import {
  parseRelationships,
  RELATIONSHIP_TYPES,
  resolveRelativePath,
} from "../../relsParser";
import {
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  type XmlElement,
} from "../../xmlParser";
import type { AnalyzedDocument } from "./blocks";
import { readPartText, relationshipsPathFor, sha256Hex } from "./package";
import { TemplatePlanError } from "./types";

/** The part a piece of content was read from. */
export type StoryOrigin = "document" | "footnotes" | "endnotes";

export const STORY_ORIGINS: readonly StoryOrigin[] = [
  "document",
  "footnotes",
  "endnotes",
];

export type StoryResources = {
  relationships: Record<StoryOrigin, RelationshipMap>;
  /**
   * The package path a relationship of `origin` targets. A target that is
   * percent-encoded or spelled in another case than its entry still resolves.
   */
  resolveTarget: (origin: StoryOrigin, target: string) => string;
  /** SHA-256 of every picture the stories embed, by package path. */
  mediaDigests: ReadonlyMap<string, string>;
  /** The bytes of those pictures, by package path. */
  media: ReadonlyMap<string, Uint8Array>;
};

const percentDecoded = (path: string): string => {
  const decoded = Result.try({
    try: () => decodeURIComponent(path),
    catch: () => null,
  });
  return Result.isError(decoded) ? path : decoded.value;
};

/** Read the relationships and embedded pictures of a package's stories. */
export const readStoryResources = async (
  analyzed: AnalyzedDocument,
): Promise<Result<StoryResources, TemplatePlanError>> => {
  const { pkg } = analyzed;
  const partPath: Record<StoryOrigin, string | null> = {
    document: pkg.documentPath,
    footnotes: analyzed.footnotes.path,
    endnotes: analyzed.endnotes.path,
  };
  const relationships: Record<StoryOrigin, RelationshipMap> = {
    document: pkg.documentRelationships,
    footnotes: new Map(),
    endnotes: new Map(),
  };
  for (const origin of ["footnotes", "endnotes"] as const) {
    const path = partPath[origin];
    const relationshipsPath = path === null ? null : relationshipsPathFor(path);
    if (relationshipsPath === null || !pkg.paths.has(relationshipsPath)) {
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const xml = await readPartText(pkg, relationshipsPath);
    if (Result.isError(xml)) {
      return xml;
    }
    relationships[origin] = parseRelationships(xml.value ?? "");
  }
  const byLowerCase = new Map(
    [...pkg.paths].map((path) => [path.toLowerCase(), path]),
  );
  const resolveTarget = (origin: StoryOrigin, target: string): string => {
    const exact = resolveRelativePath(
      relationshipsPathFor(partPath[origin] ?? pkg.documentPath),
      target,
    );
    if (pkg.paths.has(exact)) {
      return exact;
    }
    const decoded = percentDecoded(exact);
    return (
      byLowerCase.get(decoded.toLowerCase()) ??
      byLowerCase.get(exact.toLowerCase()) ??
      exact
    );
  };
  const media = new Map<string, Uint8Array>();
  const mediaDigests = new Map<string, string>();
  for (const origin of STORY_ORIGINS) {
    for (const relationship of relationships[origin].values()) {
      if (
        relationship.type !== RELATIONSHIP_TYPES.image ||
        relationship.targetMode === "External"
      ) {
        continue;
      }
      const path = resolveTarget(origin, relationship.target);
      if (media.has(path) || !pkg.paths.has(path)) {
        continue;
      }
      // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
      const bytes = await Result.tryPromise({
        try: async () => await pkg.readBytes(path),
        catch: (cause) =>
          new TemplatePlanError({
            code: "archive-unreadable",
            message: `The part ${path} could not be read.`,
            part: path,
            cause,
          }),
      });
      if (Result.isError(bytes)) {
        return bytes;
      }
      if (bytes.value !== null) {
        media.set(path, bytes.value);
        mediaDigests.set(path, sha256Hex(bytes.value));
      }
    }
  }
  return Result.ok({ relationships, resolveTarget, mediaDigests, media });
};

const isRelationshipAttribute = (
  element: XmlElement,
  name: string,
): boolean => {
  const colon = name.indexOf(":");
  return (
    colon !== -1 &&
    !name.startsWith("xmlns:") &&
    OFFICE_RELATIONSHIP_NAMESPACE_URIS.has(
      resolveNamespaceUri(element.namespaceScope, name.slice(0, colon)) ?? "",
    )
  );
};

/** The relationship ids an element and its descendants name, in document order. */
export const relationshipIdsOf = (element: XmlElement): string[] => {
  const ids: string[] = [];
  const visit = (node: XmlElement): void => {
    if (node.type !== "element") {
      return;
    }
    for (const [name, value] of Object.entries(node.attributes ?? {})) {
      if (value !== undefined && isRelationshipAttribute(node, name)) {
        ids.push(String(value));
      }
    }
    for (const child of node.elements ?? []) {
      visit(child);
    }
  };
  visit(element);
  return ids;
};

const DRAWING_PROPERTIES = "docPr";

/**
 * A digest of an object's markup that stays the same when the object moves
 * to another package: relationship ids, the drawing id and namespace
 * declarations are left out, everything else counts.
 */
export const objectFingerprint = (element: XmlElement): string => {
  const parts: string[] = [];
  const visit = (node: XmlElement): void => {
    if (node.type === "text") {
      parts.push(`#${String(node.text ?? "")}`);
      return;
    }
    if (node.type !== "element") {
      return;
    }
    const local = getLocalName(node.name);
    const attributes = Object.entries(node.attributes ?? {})
      .filter(
        ([name, value]) =>
          value !== undefined &&
          name !== "xmlns" &&
          !name.startsWith("xmlns:") &&
          !isRelationshipAttribute(node, name) &&
          !(local === DRAWING_PROPERTIES && name === "id"),
      )
      .map(([name, value]) => `${getLocalName(name)}=${String(value)}`)
      .toSorted();
    parts.push(`<${node.namespaceUri ?? ""}|${local}|${attributes.join("|")}>`);
    for (const child of node.elements ?? []) {
      visit(child);
    }
    parts.push("</>");
  };
  visit(element);
  return sha256Hex(parts.join("\u0000"));
};

/**
 * What an object's relationships lead to, as strings two packages can be
 * compared by: a picture by the digest of its bytes, an external target by
 * its address.
 */
export const objectTargets = (
  element: XmlElement,
  origin: StoryOrigin,
  resources: StoryResources,
): string[] =>
  relationshipIdsOf(element).map((id) => {
    const relationship = resources.relationships[origin].get(id);
    if (relationship === undefined) {
      return "missing";
    }
    if (relationship.targetMode === "External") {
      return `external:${relationship.type}:${relationship.target}`;
    }
    const digest = resources.mediaDigests.get(
      resources.resolveTarget(origin, relationship.target),
    );
    return digest === undefined
      ? `part:${relationship.type}`
      : `picture:${digest}`;
  });
