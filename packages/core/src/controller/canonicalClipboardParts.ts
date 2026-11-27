/**
 * The package parts copied markup names, carried from the package it was
 * copied from into the one it is pasted into.
 *
 * Kept markup (an opaque drawing, an embedded object, a kept block) names its
 * parts by relationship id: a chart, an embedding, a diagram's parts, each
 * with relationships of its own. The ids are the main document's in the body
 * and the holding part's own in a note, a comment, a header or a footer. The
 * model does not hold those parts, so a copy reads them from the source
 * package ahead of the paste, and the paste copies each one under a fresh
 * part name and relationship id.
 */

import { panic, Result } from "better-result";
import type { Slice } from "prosemirror-model";

import { escapeXmlAttribute } from "@stll/docx-core";
import { OP_STORIES, type OpStory } from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  keptMarkupIn,
  MAIN_DOCUMENT_PART,
  relationshipIdsInMarkup,
  RELATIONSHIPS_CONTENT_TYPE,
  relationshipsPartOf,
  relativeTarget,
  storyPartPath,
  storyRelationshipsKey,
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

/**
 * The kept markup of the notes and comments a slice references, read from the
 * document it was copied from, by the part holding them: `footnotes`,
 * `endnotes` or `comments`.
 */
export const sliceAnnotationMarkup = (
  slice: Slice,
  document: Document,
): Map<string, string[]> => {
  const notes = new Set<string>();
  const comments = new Set<number>();
  const commentIdOf = (attrs: Readonly<Record<string, unknown>>) => {
    const id = attrs["commentId"];
    if (typeof id === "number") comments.add(id);
  };
  slice.content.descendants((node) => {
    commentIdOf(node.attrs);
    for (const mark of node.marks) {
      commentIdOf(mark.attrs);
      const id: unknown = mark.attrs["id"];
      if (
        mark.type.name === "footnoteRef" &&
        (typeof id === "string" || typeof id === "number")
      )
        notes.add(
          `${mark.attrs["noteType"] === "endnote" ? "endnotes" : "footnotes"}:${String(id)}`,
        );
    }
    return true;
  });
  const markup = new Map<string, string[]>();
  const add = (part: string, value: unknown) => {
    const kept = keptMarkupIn(value);
    if (kept.length > 0)
      markup.set(part, [...(markup.get(part) ?? []), ...kept]);
  };
  for (const note of document.package.footnotes ?? [])
    if (notes.has(`footnotes:${String(note.id)}`)) add("footnotes", note);
  for (const note of document.package.endnotes ?? [])
    if (notes.has(`endnotes:${String(note.id)}`)) add("endnotes", note);
  for (const comment of document.package.document.comments ?? [])
    if (comments.has(comment.id)) add("comments", comment);
  return markup;
};

/** A target as a part name: percent escapes decoded where they decode. */
const decodedTarget = (target: string): string =>
  Result.try(() => decodeURIComponent(target)).unwrapOr(target);

/** The package path a relationship of the part at `partPath` targets. */
const targetOf = (partPath: string, target: string): string =>
  resolveRelativePath(relationshipsPartOf(partPath), decodedTarget(target));

/**
 * The relationships kept markup in `story` names its parts by, with the path
 * of the part holding them: the main document's for the body, the story
 * part's own otherwise.
 */
export const storyRelationshipsOf = (
  document: Document,
  story: OpStory,
): {
  relationships: ReadonlyMap<string, Relationship> | undefined;
  partPath: string | undefined;
} => {
  const key = storyRelationshipsKey(story);
  return {
    relationships:
      key === undefined
        ? document.package.relationships
        : document.package.annotationRelationships?.get(key),
    partPath: storyPartPath(document, story),
  };
};

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
 * package, with every part those parts' relationships reach: the markup
 * copied from `story` through that story's relationships, and each of
 * `annotationMarkup` through its note or comment part's own. Undefined when
 * the document has no package bytes.
 */
export const readClipboardPackageParts = async ({
  document,
  markup,
  annotationMarkup = new Map(),
  story = OP_STORIES.MAIN,
}: {
  document: Document;
  markup: readonly string[];
  /** The kept markup of copied notes and comments, by the part holding it. */
  annotationMarkup?: ReadonlyMap<string, readonly string[]>;
  /** The story the markup was copied from; the body when absent. */
  story?: OpStory;
}): Promise<ClipboardPackageParts | undefined> => {
  // A part an earlier paste copied in lives in `copiedParts` until a save
  // writes it; it reads from there before the package it was opened from.
  const copied = new Map(
    [...(document.package.copiedParts?.values() ?? [])].map((part) => [
      part.path.toLowerCase(),
      part,
    ]),
  );
  const buffer = document.originalBuffer;
  if (buffer === undefined && copied.size === 0) return undefined;
  const rootsOf = (
    kept: readonly string[],
    { relationships, partPath }: ReturnType<typeof storyRelationshipsOf>,
  ): string[] => {
    if (partPath === undefined) return [];
    const ids = new Set(
      kept.flatMap((xml) => relationshipIdsInMarkup(xml) ?? []),
    );
    return [...ids].flatMap((id) => {
      const relationship = relationships?.get(id);
      return relationship === undefined ||
        relationship.targetMode === "External"
        ? []
        : [targetOf(partPath, relationship.target)];
    });
  };
  // Note and comment parts sit beside the main document.
  const roots = [
    ...rootsOf(markup, storyRelationshipsOf(document, story)),
    ...[...annotationMarkup].flatMap(([part, kept]) =>
      // A part without its own relationships resolves through the main
      // document's, as the parser reads it.
      rootsOf(kept, {
        relationships:
          document.package.annotationRelationships?.get(part) ??
          document.package.relationships,
        partPath: `word/${part}.xml`,
      }),
    ),
  ];
  // Markup naming only external targets copies its relationships alone.
  if (roots.length === 0) return { source: document, parts: new Map() };
  const zip = buffer === undefined ? undefined : await ZipArchive.read(buffer);
  const entries = new Map(
    (zip?.entries() ?? []).map((entry) => [entry.path.toLowerCase(), entry]),
  );
  const registry = entries.get("[content_types].xml");
  const contentTypeOf = contentTypesOf(
    registry === undefined
      ? ""
      : decodeText(await registry.bytes(), { ignoreBOM: true }),
  );
  const partAt = async (path: string): Promise<SourcePart | undefined> => {
    const held = copied.get(path.toLowerCase());
    if (held !== undefined)
      return { data: new Uint8Array(held.data), contentType: held.contentType };
    const entry = entries.get(path.toLowerCase());
    const contentType = contentTypeOf(path);
    if (entry === undefined || contentType === undefined) return undefined;
    return { data: await entry.bytes(), contentType };
  };
  const parts = new Map<string, SourcePart>();
  const pending = [...roots];
  for (let path = pending.pop(); path !== undefined; path = pending.pop()) {
    if (parts.has(path)) continue;
    // oxlint-disable-next-line no-await-in-loop -- each part names the next
    const part = await partAt(path);
    if (part === undefined) continue;
    parts.set(path, part);
    const relsPath = relationshipsPartOf(path);
    const heldRels = copied.get(relsPath.toLowerCase());
    const relsEntry = entries.get(relsPath.toLowerCase());
    let data: Uint8Array;
    if (heldRels !== undefined) data = new Uint8Array(heldRels.data);
    else if (relsEntry !== undefined)
      // oxlint-disable-next-line no-await-in-loop -- each part names the next
      data = await relsEntry.bytes();
    else continue;
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
  /**
   * The source relationships the markup's ids name: the main document's, or
   * those of the story part the markup was copied from.
   */
  sourceRelationships: ReadonlyMap<string, Relationship>;
  /** The part `sourceRelationships` belong to; the main document when absent. */
  sourcePart?: string;
  /** The paste's destination, whose package the copies must not collide with. */
  destination: Document;
  /**
   * The destination's relationships, which receive the copied ones: the main
   * document's, or those of the story part the markup lands in.
   */
  relationships: Map<string, Relationship>;
  /** The part `relationships` belong to; the main document when absent. */
  destinationPart?: string;
  /** The destination's copied parts, which receive the copies. */
  copiedParts: Map<string, CopiedPackagePart>;
  /** A relationship id the destination does not hold yet. */
  freshRelationshipId: () => string;
  /**
   * The name each source part was copied under, shared by the copiers of one
   * paste so a part the body's, a note's and a comment's markup all name is
   * copied once.
   */
  copiedPaths?: Map<string, string>;
};

/**
 * Why a copy failed: the source relationships do not hold what the markup
 * names (`unresolved`), or a part it reaches was not read from the source
 * package or the destination package cannot be read (`unreadable`).
 */
export type PartCopyFailure = { type: "unresolved" | "unreadable" };

/** The folder every copied part lives under, beside the parts naming it. */
const COPY_FOLDER = "word/";

/**
 * Copies the parts kept markup names into the destination: each relationship
 * under a fresh id, each internal target with its own relationships under a
 * fresh part name beside its source name. A part outside the main document's
 * folder is copied into it, keeping its own folder below it. A part two
 * pieces of markup share is copied once, and a copy that cannot reach every
 * part it needs copies nothing.
 */
export const createPartCopier = ({
  parts,
  sourceRelationships,
  sourcePart = MAIN_DOCUMENT_PART,
  destination,
  relationships,
  destinationPart = MAIN_DOCUMENT_PART,
  copiedParts,
  freshRelationshipId,
  copiedPaths = new Map(),
}: PartCopierOptions) => {
  const copiedIds = new Map<string, string>();
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
        names.add(targetOf(destinationPart, target).toLowerCase());
    taken = names;
    return Result.ok(names);
  };
  /**
   * A part name beside `path` that the destination does not hold, under the
   * main document's folder.
   */
  const freshName = (path: string, names: Set<string>): string => {
    const placed = path.startsWith(COPY_FOLDER)
      ? path
      : `${COPY_FOLDER}${path}`;
    const slash = placed.lastIndexOf("/");
    const folder = placed.slice(0, slash + 1);
    const name = placed.slice(slash + 1);
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
  /** The relationships a read part names, with the part holding them. */
  const partRelationships = (path: string) => {
    const relsPath = relationshipsPartOf(path);
    const rels = parts.parts.get(relsPath);
    return rels === undefined
      ? undefined
      : {
          relsPath,
          relationships: [
            ...parseRelationships(
              decodeText(rels.data, { ignoreBOM: true }),
            ).values(),
          ],
        };
  };
  /** Whether every part `path` reaches was read, so it can be copied whole. */
  const reachable = (path: string, seen: Set<string>): boolean => {
    if (seen.has(path) || copiedPaths.has(path)) return true;
    seen.add(path);
    if (!parts.parts.has(path)) return false;
    const named = partRelationships(path);
    return (
      named === undefined ||
      named.relationships.every(
        (relationship) =>
          relationship.targetMode === "External" ||
          reachable(
            resolveRelativePath(
              named.relsPath,
              decodedTarget(relationship.target),
            ),
            seen,
          ),
      )
    );
  };
  const copyPart = (path: string, names: Set<string>): string => {
    const copied = copiedPaths.get(path);
    if (copied !== undefined) return copied;
    const part =
      parts.parts.get(path) ?? panic("A reachable copied part was not read.");
    const target = freshName(path, names);
    names.add(target.toLowerCase());
    copiedPaths.set(path, target);
    const named = partRelationships(path);
    if (named !== undefined) {
      const entries: string[] = [];
      for (const relationship of named.relationships) {
        const relationshipTarget =
          relationship.targetMode === "External"
            ? relationship.target
            : relativeTarget(
                target,
                copyPart(
                  resolveRelativePath(
                    named.relsPath,
                    decodedTarget(relationship.target),
                  ),
                  names,
                ),
              );
        entries.push(
          `<Relationship Id="${escapeXmlAttribute(relationship.id)}" Type="${escapeXmlAttribute(relationship.type)}" Target="${escapeXmlAttribute(relationshipTarget)}"${relationship.targetMode === "External" ? ' TargetMode="External"' : ""}/>`,
        );
      }
      const copiedRels = relationshipsPartOf(target);
      names.add(copiedRels.toLowerCase());
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
    return target;
  };
  /** The destination relationship id the copy of `id` takes. */
  const copyRelationship = (id: string, names: Set<string>): string => {
    const copied = copiedIds.get(id);
    if (copied !== undefined) return copied;
    const source =
      sourceRelationships.get(id) ??
      panic("A checked copied relationship is missing.");
    const target =
      source.targetMode === "External"
        ? source.target
        : relativeTarget(
            destinationPart,
            copyPart(targetOf(sourcePart, source.target), names),
          );
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
    return next;
  };
  return {
    /** Copy what each of `ids` names; the map from each to its copy's id. */
    copy: (
      ids: readonly string[],
    ): Result<ReadonlyMap<string, string>, PartCopyFailure> => {
      // Every id resolves and every part it reaches was read before any is
      // copied, so a copy that fails leaves the destination as it was.
      const seen = new Set<string>();
      for (const id of ids) {
        const source = sourceRelationships.get(id);
        if (source === undefined) return Result.err({ type: "unresolved" });
        if (
          source.targetMode !== "External" &&
          !copiedIds.has(id) &&
          !reachable(targetOf(sourcePart, source.target), seen)
        )
          return Result.err({ type: "unreadable" });
      }
      const names = takenNames();
      if (names.isErr()) return names;
      const renamed = new Map<string, string>();
      for (const id of ids) renamed.set(id, copyRelationship(id, names.value));
      return Result.ok(renamed);
    },
  };
};
