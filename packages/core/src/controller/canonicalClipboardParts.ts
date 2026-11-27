/**
 * The package parts copied markup names, carried from the package it was
 * copied from into the one it is pasted into.
 *
 * Kept markup (an opaque drawing, an embedded object, a kept block) names its
 * parts by relationship id: a chart, an embedding, a diagram's parts, each
 * with relationships of its own. The model does not hold those parts, so a
 * copy reads them from the source package ahead of the paste, and the paste
 * copies each one under a fresh part name and relationship id.
 */

import { Result } from "better-result";
import type { Slice } from "prosemirror-model";

import { escapeXmlAttribute } from "@stll/docx-core";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  mainRelationshipTarget,
  relationshipIdsInMarkup,
  RELATIONSHIPS_CONTENT_TYPE,
  relationshipsPartOf,
  relativeTarget,
} from "../docx/copiedPackageParts";
import { parseRelationships, resolveRelativePath } from "../docx/relsParser";
import {
  getAttribute,
  getChildElements,
  getLocalName,
  parseXmlDocument,
} from "../docx/xmlParser";
import type {
  CopiedPackagePart,
  Document,
  Relationship,
} from "../types/document";
import { decodeText, encodeUtf8 } from "../utils/textEncoding";

/** A part read from the source package. */
type SourcePart = { readonly data: Uint8Array; readonly contentType: string };

/**
 * The parts of a copy's source package that its kept markup reaches, read
 * ahead of a paste into another package, by their paths in that package.
 */
export type ClipboardPackageParts = {
  /** The document the parts were read from: a paste uses them only with it. */
  readonly source: Document;
  readonly parts: ReadonlyMap<string, SourcePart>;
};

/** Every piece of markup a slice's nodes keep. */
export const sliceKeptMarkup = (slice: Slice): string[] => {
  const markup: string[] = [];
  slice.content.descendants((node) => {
    for (const value of Object.values(node.attrs))
      if (typeof value === "string" && value.startsWith("<"))
        markup.push(value);
    return true;
  });
  return markup;
};

/** A target as a part name: percent escapes decoded where they decode. */
const decodedTarget = (target: string): string =>
  Result.try(() => decodeURIComponent(target)).unwrapOr(target);

/** The content type the registry gives each part, by lower-cased path. */
const contentTypesOf = (xml: string) => {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const element of getChildElements(parseXmlDocument(xml))) {
    const type = getAttribute(element, null, "ContentType");
    if (type === null) continue;
    const name = getLocalName(element.name ?? "");
    const partName = getAttribute(element, null, "PartName");
    const extension = getAttribute(element, null, "Extension");
    if (name === "Override" && partName !== null)
      overrides.set(
        decodedTarget(partName).replace(/^\//u, "").toLowerCase(),
        type,
      );
    if (name === "Default" && extension !== null)
      defaults.set(extension.toLowerCase(), type);
  }
  return (path: string): string | undefined =>
    overrides.get(path.toLowerCase()) ??
    defaults.get(path.slice(path.lastIndexOf(".") + 1).toLowerCase());
};

/**
 * Read the parts the kept markup among `markup` names from `document`'s own
 * package, with every part those parts' relationships reach. Undefined when
 * the markup names no relationship or the document has no package bytes.
 */
export const readClipboardPackageParts = async ({
  document,
  markup,
}: {
  document: Document;
  markup: readonly string[];
}): Promise<ClipboardPackageParts | undefined> => {
  const buffer = document.originalBuffer;
  const relationships = document.package.relationships;
  if (buffer === undefined || relationships === undefined) return undefined;
  const ids = new Set(
    markup.flatMap((xml) => relationshipIdsInMarkup(xml) ?? []),
  );
  const roots = [...ids].flatMap((id) => {
    const relationship = relationships.get(id);
    return relationship === undefined || relationship.targetMode === "External"
      ? []
      : [mainRelationshipTarget(decodedTarget(relationship.target))];
  });
  if (roots.length === 0) return undefined;
  const zip = await ZipArchive.read(buffer);
  const entries = new Map(
    zip.entries().map((entry) => [entry.path.toLowerCase(), entry]),
  );
  const registry = entries.get("[content_types].xml");
  const contentTypeOf = contentTypesOf(
    registry === undefined
      ? ""
      : decodeText(await registry.bytes(), { ignoreBOM: true }),
  );
  const parts = new Map<string, SourcePart>();
  const pending = [...roots];
  for (let path = pending.pop(); path !== undefined; path = pending.pop()) {
    if (parts.has(path)) continue;
    const entry = entries.get(path.toLowerCase());
    const contentType = contentTypeOf(path);
    if (entry === undefined || contentType === undefined) continue;
    // oxlint-disable-next-line no-await-in-loop -- each part names the next
    parts.set(path, { data: await entry.bytes(), contentType });
    const relsPath = relationshipsPartOf(path);
    const rels = entries.get(relsPath.toLowerCase());
    if (rels === undefined) continue;
    // oxlint-disable-next-line no-await-in-loop -- each part names the next
    const data = await rels.bytes();
    parts.set(relsPath, { data, contentType: RELATIONSHIPS_CONTENT_TYPE });
    for (const relationship of parseRelationships(
      decodeText(data, { ignoreBOM: true }),
    ).values())
      if (relationship.targetMode !== "External")
        pending.push(
          resolveRelativePath(relsPath, decodedTarget(relationship.target)),
        );
  }
  return { source: document, parts };
};

type PartCopierOptions = {
  parts: ClipboardPackageParts;
  /** The source document's main relationships, which the markup's ids name. */
  sourceRelationships: ReadonlyMap<string, Relationship>;
  /** The paste's destination, whose package the copies must not collide with. */
  destination: Document;
  /** The destination's relationships, which receive the copied ones. */
  relationships: Map<string, Relationship>;
  /** The destination's copied parts, which receive the copies. */
  copiedParts: Map<string, CopiedPackagePart>;
  /** A relationship id the destination does not hold yet. */
  freshRelationshipId: () => string;
};

/** Why a copy failed: the source does not hold what the markup names. */
export type PartCopyFailure = { type: "unresolved" | "unreadable" };

/**
 * Copies the parts kept markup names into the destination: each relationship
 * under a fresh id, each internal target with its own relationships under a
 * fresh part name beside its source name. A part two pieces of markup share
 * is copied once.
 */
export const createPartCopier = ({
  parts,
  sourceRelationships,
  destination,
  relationships,
  copiedParts,
  freshRelationshipId,
}: PartCopierOptions) => {
  const copiedIds = new Map<string, string>();
  const copiedPaths = new Map<string, string>();
  let taken: Set<string> | undefined;
  /** Every part name the destination holds or has been given, lower-cased. */
  const takenNames = (): Result<Set<string>, PartCopyFailure> => {
    if (taken !== undefined) return Result.ok(taken);
    const names = new Set<string>();
    if (destination.originalBuffer !== undefined) {
      const zip = ZipArchive.readDirectory(destination.originalBuffer);
      if (zip.isErr()) return Result.err({ type: "unreadable" });
      for (const path of zip.value.paths()) names.add(path.toLowerCase());
    }
    for (const path of destination.package.media?.keys() ?? [])
      names.add(path.toLowerCase());
    for (const path of copiedParts.keys()) names.add(path.toLowerCase());
    for (const { target, targetMode } of relationships.values())
      if (targetMode !== "External")
        names.add(mainRelationshipTarget(target).toLowerCase());
    taken = names;
    return Result.ok(names);
  };
  /** A part name beside `path` that the destination does not hold. */
  const freshName = (path: string, names: Set<string>): string => {
    const slash = path.lastIndexOf("/");
    const folder = path.slice(0, slash + 1);
    const name = path.slice(slash + 1);
    const dot = name.lastIndexOf(".");
    const extension = dot > 0 ? name.slice(dot) : "";
    const stem = (dot > 0 ? name.slice(0, dot) : name).replace(/\d+$/u, "");
    for (let ordinal = 1; ; ordinal += 1) {
      const candidate = `${folder}${stem}${ordinal}${extension}`;
      if (
        !names.has(candidate.toLowerCase()) &&
        !names.has(relationshipsPartOf(candidate).toLowerCase())
      )
        return candidate;
    }
  };
  const copyPart = (path: string): Result<string, PartCopyFailure> => {
    const copied = copiedPaths.get(path);
    if (copied !== undefined) return Result.ok(copied);
    const part = parts.parts.get(path);
    if (part === undefined || !path.startsWith("word/"))
      return Result.err({ type: "unreadable" });
    const names = takenNames();
    if (names.isErr()) return names;
    const target = freshName(path, names.value);
    names.value.add(target.toLowerCase());
    copiedPaths.set(path, target);
    const relsPath = relationshipsPartOf(path);
    const rels = parts.parts.get(relsPath);
    if (rels !== undefined) {
      const entries: string[] = [];
      for (const relationship of parseRelationships(
        decodeText(rels.data, { ignoreBOM: true }),
      ).values()) {
        let relationshipTarget = relationship.target;
        if (relationship.targetMode !== "External") {
          const copiedTarget = copyPart(
            resolveRelativePath(relsPath, decodedTarget(relationship.target)),
          );
          if (copiedTarget.isErr()) return copiedTarget;
          relationshipTarget = relativeTarget(target, copiedTarget.value);
        }
        entries.push(
          `<Relationship Id="${escapeXmlAttribute(relationship.id)}" Type="${escapeXmlAttribute(relationship.type)}" Target="${escapeXmlAttribute(relationshipTarget)}"${relationship.targetMode === "External" ? ' TargetMode="External"' : ""}/>`,
        );
      }
      const copiedRels = relationshipsPartOf(target);
      names.value.add(copiedRels.toLowerCase());
      copiedParts.set(copiedRels, {
        path: copiedRels,
        contentType: RELATIONSHIPS_CONTENT_TYPE,
        data: encodeUtf8(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join("")}</Relationships>`,
        ).buffer,
      });
    }
    copiedParts.set(target, {
      path: target,
      contentType: part.contentType,
      data: new Uint8Array(part.data).buffer,
    });
    return Result.ok(target);
  };
  /** The destination relationship id the copy of `id` takes. */
  const copyRelationship = (id: string): Result<string, PartCopyFailure> => {
    const copied = copiedIds.get(id);
    if (copied !== undefined) return Result.ok(copied);
    const source = sourceRelationships.get(id);
    if (source === undefined) return Result.err({ type: "unresolved" });
    let target = source.target;
    if (source.targetMode !== "External") {
      const part = copyPart(
        mainRelationshipTarget(decodedTarget(source.target)),
      );
      if (part.isErr()) return part;
      target = relativeTarget("word/document.xml", part.value);
    }
    const next = freshRelationshipId();
    relationships.set(next, {
      id: next,
      type: source.type,
      target,
      ...(source.targetMode === undefined
        ? {}
        : { targetMode: source.targetMode }),
    });
    copiedIds.set(id, next);
    return Result.ok(next);
  };
  return {
    /** Copy what each of `ids` names; the map from each to its copy's id. */
    copy: (
      ids: readonly string[],
    ): Result<ReadonlyMap<string, string>, PartCopyFailure> => {
      const renamed = new Map<string, string>();
      for (const id of ids) {
        const copied = copyRelationship(id);
        if (copied.isErr()) return copied;
        renamed.set(id, copied.value);
      }
      return Result.ok(renamed);
    },
  };
};
