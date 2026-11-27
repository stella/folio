/**
 * Content placed into a part of the package other than the one it was read
 * from, with every relationship it names re-homed into the part it lands in.
 *
 * A picture's `r:embed`, a link's `r:id`, a drawing's or an embedded object's
 * kept markup, and kept block markup name relationships of the part that holds
 * them: the main document's for the body (and a text box in it), a header's or
 * a footer's own, and the notes' and comments' parts' own. The same id names
 * something else in another part, or nothing. So content moved or copied from
 * one part into another names, in the part it lands in, an id that names there
 * what it named where it was read: the id itself where the destination already
 * holds it for the same target, another id the destination holds for that
 * target, or otherwise a fresh id the destination is given for it. The content
 * is rewritten to those ids, and one package-resources operation gives the
 * destination the relationships it lacked, which the save writes and the
 * operation's inverse takes back.
 */

import { panic } from "better-result";

import { relationshipIdOf } from "@stll/docx-core/model";
import {
  OP_STORIES,
  type OpStory,
  packageResourcesOpOf,
  type SetPackageResourcesOp,
} from "@stll/docx-core/ops";

import type { Document, Image, Relationship } from "../types/document";
import {
  isHeaderFooterRelationshipsKey,
  KEPT_MARKUP_KEYS,
  partRelationshipsOf,
  relationshipIdsInMarkup,
  relationshipTargetPath,
  relativeTarget,
  sameRelationship,
  storyPartPath,
  storyRelationshipsKey,
  withRelationshipIds,
} from "./copiedPackageParts";
import { imageRawXmlFingerprint } from "./imageRawXml";

/** The comments part, which holds every comment's content. */
export const COMMENTS_PART = "comments";

/**
 * Where content is read from or placed: a story, or the comments part. Each
 * names the part whose relationships the content's ids resolve in.
 */
export type RelationshipStory = OpStory | typeof COMMENTS_PART;

/** The key the part's own relationships are kept under; undefined for the main document. */
const partKeyOf = (story: RelationshipStory): string | undefined =>
  story === COMMENTS_PART ? COMMENTS_PART : storyRelationshipsKey(story);

/**
 * The path of the part holding `story`. A header or footer the save has not
 * written yet is written beside the main document, under `word/`.
 */
const partPathOf = (document: Document, story: RelationshipStory): string => {
  if (story === COMMENTS_PART) return `word/${COMMENTS_PART}.xml`;
  const path = storyPartPath(document, story);
  if (path !== undefined) return path;
  if (story === OP_STORIES.MAIN || !("rId" in story))
    panic("Every part besides a header or a footer has a fixed path.");
  return `word/${story.kind}.xml`;
};

/** Whether content read from `one` names its relationships as `other` does. */
export const sharesRelationships = (
  one: RelationshipStory,
  other: RelationshipStory,
): boolean => partKeyOf(one) === partKeyOf(other);

/** The records whose `rId` names a relationship of the part holding them. */
const RELATIONSHIP_RECORDS: ReadonlySet<unknown> = new Set([
  "hyperlink",
  "image",
]);

type Renaming = {
  /** The id a modelled reference takes; undefined keeps it. */
  id: (id: string) => string | undefined;
  /** Kept markup with the ids it names taken. */
  markup: (xml: string) => string;
};

/** The picture of a drawing whose captured markup carries a fingerprint of it. */
const fingerprintedPicture = (value: object): Image | undefined => {
  if (
    Reflect.get(value, "type") !== "drawing" ||
    Reflect.get(value, "rawImageFingerprint") === undefined
  )
    return undefined;
  const image: unknown = Reflect.get(value, "image");
  return typeof image === "object" &&
    image !== null &&
    Reflect.get(image, "type") === "image"
    ? (image as Image)
    : undefined;
};

/**
 * Take every relationship reference below `value`, in place: a link's and a
 * picture's ids, an image's click and hover links, and the ids kept markup
 * names. A section's header and footer references name the main document's
 * relationships wherever the section is, so they stay.
 */
const renameReferences = (
  value: unknown,
  renaming: Renaming,
  linkSource = false,
): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) renameReferences(item, renaming);
    return;
  }
  // A drawing whose captured markup still agrees with its picture keeps
  // agreeing once both take the same ids.
  const picture = fingerprintedPicture(value);
  const current =
    picture !== undefined &&
    Reflect.get(value, "rawImageFingerprint") ===
      imageRawXmlFingerprint(picture);
  const named =
    linkSource || RELATIONSHIP_RECORDS.has(Reflect.get(value, "type"));
  for (const [key, item] of Object.entries(value)) {
    if (key === "sectionProperties") continue;
    if (typeof item === "string") {
      if (KEPT_MARKUP_KEYS.has(key) && item.startsWith("<"))
        Reflect.set(value, key, renaming.markup(item));
      else if (named && (key === "rId" || key === "hlinkRId")) {
        const next = renaming.id(item);
        if (next !== undefined) Reflect.set(value, key, next);
      }
      continue;
    }
    renameReferences(item, renaming, key === "hlinkClickSource");
  }
  if (picture !== undefined && current)
    Reflect.set(value, "rawImageFingerprint", imageRawXmlFingerprint(picture));
};

/** Every relationship id `value` names, modelled or in its kept markup. */
const referencedIds = (value: unknown): Set<string> => {
  const ids = new Set<string>();
  renameReferences(structuredClone(value), {
    id: (id) => {
      ids.add(id);
      return undefined;
    },
    markup: (xml) => {
      for (const id of relationshipIdsInMarkup(xml) ?? []) ids.add(id);
      return xml;
    },
  });
  return ids;
};

export type RelationshipRehomer = {
  /**
   * A copy of `content` read from `from`'s part, naming in `to`'s part what
   * it named there; `content` itself where the two parts are one.
   */
  rehome: <Content>(
    content: Content,
    placement: { from: RelationshipStory; to: RelationshipStory },
  ) => Content;
  /**
   * The operation giving each destination part the relationships the
   * re-homed content names that it lacked; undefined when none was added.
   */
  resourcesOp: () => SetPackageResourcesOp | undefined;
};

/** Re-homes content placed into the parts of `document`. */
export const createRelationshipRehomer = (
  document: Document,
): RelationshipRehomer => {
  const pkg = document.package;
  let main: Map<string, Relationship> | undefined;
  let mainChanged = false;
  const parts = new Map<string, Map<string, Relationship>>();
  const changedParts = new Set<string>();
  const mainTable = () => {
    main ??= new Map(pkg.relationships);
    return main;
  };
  /**
   * The relationships content lands in, the rehomer's own copy. A note or
   * comment part without its own resolves its ids through the main
   * document's, so the table it is given starts from those: the ids its
   * content already names keep naming what they named.
   */
  const destinationTable = (key: string | undefined) => {
    if (key === undefined) return mainTable();
    const held = parts.get(key);
    if (held !== undefined) return held;
    const own = pkg.annotationRelationships?.get(key);
    const table = new Map(own);
    if (own === undefined && !isHeaderFooterRelationshipsKey(key))
      for (const [id, relationship] of pkg.relationships ?? [])
        table.set(id, relationship);
    parts.set(key, table);
    return table;
  };
  /**
   * An id no part holds: a picture pasted into a part takes a main document
   * id that the save writes into that part, so an id is fresh only where no
   * table, and no header or footer awaiting its relationship, holds it.
   */
  let ordinal = 1;
  const held = (id: string) =>
    mainTable().has(id) ||
    pkg.headers?.has(id) === true ||
    pkg.footers?.has(id) === true ||
    [...parts.values()].some((table) => table.has(id)) ||
    [...(pkg.annotationRelationships?.values() ?? [])].some((table) =>
      table.has(id),
    );
  const freshId = () => {
    while (held(`rId${ordinal}`)) ordinal += 1;
    const id = `rId${ordinal}`;
    ordinal += 1;
    return id;
  };
  return {
    rehome: (content, { from, to }) => {
      const fromKey = partKeyOf(from);
      const toKey = partKeyOf(to);
      if (fromKey === toKey) return content;
      const ids = referencedIds(content);
      if (ids.size === 0) return content;
      const fromPath = partPathOf(document, from);
      const toPath = partPathOf(document, to);
      const source =
        fromKey === undefined
          ? pkg.relationships
          : partRelationshipsOf(document, fromKey);
      const into = destinationTable(toKey);
      const renamed = new Map<string, string>();
      for (const id of ids) {
        const relationship = source?.get(id);
        // An id its own part does not hold names nothing anywhere.
        if (relationship === undefined) continue;
        const own = { relationship, part: fromPath };
        if (sameRelationship(own, { relationship: into.get(id), part: toPath }))
          continue;
        const reused = [...into.values()].find((candidate) =>
          sameRelationship(own, { relationship: candidate, part: toPath }),
        );
        if (reused !== undefined) {
          renamed.set(id, reused.id);
          continue;
        }
        const next = freshId();
        into.set(next, {
          id: next,
          type: relationship.type,
          target:
            relationship.targetMode === "External"
              ? relationship.target
              : relativeTarget(
                  toPath,
                  relationshipTargetPath(fromPath, relationship.target),
                ),
          ...(relationship.targetMode === undefined
            ? {}
            : { targetMode: relationship.targetMode }),
        });
        if (toKey === undefined) mainChanged = true;
        else changedParts.add(toKey);
        renamed.set(id, next);
      }
      if (renamed.size === 0) return content;
      const placed = structuredClone(content);
      renameReferences(placed, {
        id: (id) => {
          const next = renamed.get(id);
          return next === undefined ? undefined : relationshipIdOf(next);
        },
        markup: (xml) => withRelationshipIds(xml, renamed) ?? xml,
      });
      return placed;
    },
    resourcesOp: () => {
      if (!mainChanged && changedParts.size === 0) return undefined;
      const tables = new Map(pkg.annotationRelationships);
      for (const key of changedParts)
        tables.set(
          key,
          parts.get(key) ?? panic("A changed part has no relationships."),
        );
      const after: Document = {
        ...document,
        package: {
          ...pkg,
          ...(mainChanged ? { relationships: mainTable() } : {}),
          ...(changedParts.size > 0 ? { annotationRelationships: tables } : {}),
        },
      };
      return packageResourcesOpOf({ before: document, after });
    },
  };
};
