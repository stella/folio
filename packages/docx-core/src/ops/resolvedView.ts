/**
 * The accepted or the rejected view of a document: every revision of every
 * story resolved one way, through the same resolution an Accept All or a
 * Reject All applies. Readers that show a document as if its changes were
 * accepted or rejected read it from here, so they and a reviewer resolving
 * the document agree on every paragraph, table and note it keeps.
 */

import type { Document } from "../model/document";
import { applyDocumentOps } from "./apply";
import { IDENTITY_SPACES, identityKeysIn, packageParagraphIds } from "./ids";
import { documentStories, findStoryBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type OpStory,
  type RevisionDecision,
} from "./types";

/** A document with its changes resolved one way, and what could not be. */
export type ResolvedDocumentView = {
  document: Document;
  /**
   * The stories whose resolution refused, left as they were: a reader shows
   * them with their changes still in place and says so.
   */
  unresolved: readonly OpStory[];
};

const REVISION_KEY_PREFIX = `${IDENTITY_SPACES.REVISION}:`;

/** Every revision id a story's content carries, in document order. */
const storyRevisionIds = (document: Document, story: OpStory): number[] => {
  const body = findStoryBody(document, story);
  if (body === undefined) return [];
  return [
    ...new Set(
      identityKeysIn(body.content).flatMap((key) =>
        key.startsWith(REVISION_KEY_PREFIX)
          ? [Number(key.slice(REVISION_KEY_PREFIX.length))]
          : [],
      ),
    ),
  ];
};

const isPlainRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" &&
  value !== null &&
  Object.getPrototypeOf(value) === Object.prototype;

/**
 * `value` with `edit` applied to every plain record, bottom up, copying only
 * what changes. A record reached twice (a section's view of the body shares
 * the body's paragraphs) maps to one copy, except one `each` names: that one
 * maps anew in every place.
 */
const mapRecords = (
  value: unknown,
  edit: (record: Record<string, unknown>) => Record<string, unknown>,
  each: (record: Record<string, unknown>) => boolean = () => false,
  seen: Map<object, unknown> = new Map(),
): unknown => {
  if (typeof value !== "object" || value === null) return value;
  const fresh = isPlainRecord(value) && each(value);
  const known = fresh ? undefined : seen.get(value);
  if (known !== undefined) return known;
  let out: unknown = value;
  if (Array.isArray(value)) {
    const items = value.map((item) => mapRecords(item, edit, each, seen));
    if (items.some((item, index) => item !== value[index])) out = items;
  } else if (value instanceof Map) {
    const entries = [...value.entries()].map(
      ([key, item]) => [key, mapRecords(item, edit, each, seen)] as const,
    );
    if (entries.some(([key, item]) => item !== value.get(key)))
      out = new Map(entries);
  } else if (isPlainRecord(value)) {
    let record = value;
    for (const [key, item] of Object.entries(value)) {
      const mapped = mapRecords(item, edit, each, seen);
      if (mapped === item) continue;
      if (record === value) record = { ...value };
      record[key] = mapped;
    }
    out = edit(record);
  }
  if (!fresh) seen.set(value, out);
  return out;
};

/**
 * `document` with an id on every paragraph, which resolution addresses
 * paragraphs by, and the ids it minted: the view leaves them off again.
 */
const withParagraphIds = (
  document: Document,
): { document: Document; minted: ReadonlySet<string> } => {
  const taken = new Set(packageParagraphIds(document.package));
  const minted = new Set<string>();
  let next = 0x10000000;
  const fresh = (): string => {
    let id = next.toString(16).toUpperCase();
    while (taken.has(id)) {
      next += 1;
      id = next.toString(16).toUpperCase();
    }
    next += 1;
    minted.add(id);
    return id;
  };
  const pkg = mapRecords(document.package, (record) =>
    record["type"] === "paragraph" && typeof record["paraId"] !== "string"
      ? { ...record, paraId: fresh() }
      : record,
  ) as Document["package"];
  return {
    document:
      pkg === document.package ? document : { ...document, package: pkg },
    minted,
  };
};

/**
 * `document` with an id of its own on every revision record. A host model can
 * let records share one (an inserted paragraph's mark and its words), as its
 * writer later numbers them apart, where an operation refuses the repeat; the
 * view resolves every record anyway, so it numbers them apart first.
 */
const revisionIdOf = (record: Record<string, unknown>): number | undefined => {
  const info = record["info"];
  return isPlainRecord(info) &&
    typeof info["id"] === "number" &&
    typeof info["author"] === "string"
    ? info["id"]
    : undefined;
};

const isRevisionRecord = (record: Record<string, unknown>): boolean =>
  revisionIdOf(record) !== undefined;

const withDistinctRevisionIds = (document: Document): Document => {
  // Every place a revision record stands counts and is numbered on its own,
  // one object a model shares between paragraphs included: its writer gives
  // each place its own id, and the identity census reads equal shared records
  // as one. A paragraph shared whole (a section's view) stays one place.
  const ids: number[] = [];
  mapRecords(
    document.package,
    (record) => {
      const id = revisionIdOf(record);
      if (id !== undefined) ids.push(id);
      return record;
    },
    isRevisionRecord,
  );
  if (new Set(ids).size === ids.length) return document;
  let next = Math.max(...ids) + 1;
  const seen = new Set<number>();
  const pkg = mapRecords(
    document.package,
    (record) => {
      const id = revisionIdOf(record);
      if (id === undefined) return record;
      if (!seen.has(id)) {
        seen.add(id);
        return record;
      }
      const info = record["info"] as Record<string, unknown>;
      return { ...record, info: { ...info, id: next++ } };
    },
    isRevisionRecord,
  ) as Document["package"];
  return pkg === document.package ? document : { ...document, package: pkg };
};

const withoutParagraphIds = (
  document: Document,
  minted: ReadonlySet<string>,
): Document => {
  if (minted.size === 0) return document;
  const pkg = mapRecords(document.package, (record) => {
    const id = record["paraId"];
    if (
      record["type"] !== "paragraph" ||
      typeof id !== "string" ||
      !minted.has(id)
    )
      return record;
    return Object.fromEntries(
      Object.entries(record).filter(([key]) => key !== "paraId"),
    );
  }) as Document["package"];
  return pkg === document.package ? document : { ...document, package: pkg };
};

/**
 * `document` with every revision resolved by `decision`, story by story. A
 * story the resolution of an earlier one removed (a note whose every
 * reference went) is not read again.
 */
export const resolvedDocumentView = (
  document: Document,
  decision: RevisionDecision,
): ResolvedDocumentView => {
  const stories = documentStories(document);
  if (stories.every((story) => storyRevisionIds(document, story).length === 0))
    return { document, unresolved: [] };
  const identified = withParagraphIds(withDistinctRevisionIds(document));
  let current = identified.document;
  const unresolved: OpStory[] = [];
  for (const story of stories) {
    const revisionIds = storyRevisionIds(current, story);
    if (revisionIds.length === 0) continue;
    const resolved = applyDocumentOps(current, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds,
        decision,
      },
    ]);
    if (resolved.isErr()) unresolved.push(story);
    else current = resolved.value.document;
  }
  return {
    document: withoutParagraphIds(current, identified.minted),
    unresolved,
  };
};
