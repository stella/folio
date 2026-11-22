/**
 * The accepted or the rejected view of a document: every revision of every
 * story resolved one way, through the same resolution an Accept All or a
 * Reject All applies. Readers that show a document as if its changes were
 * accepted or rejected read it from here, so they and a reviewer resolving
 * the document agree on every paragraph, table and note it keeps.
 */

import {
  type BlockContent,
  type Document,
  REVIEW_CARRIERS,
} from "../model/document";
import { maxOf } from "../model/extremes";
import { applyDocumentOps } from "./apply";
import { withBodyContent } from "./blocks";
import { IDENTITY_SPACES, identityKeysIn, packageParagraphIds } from "./ids";
import {
  documentStories,
  findStoryBody,
  replaceStoryBody,
  storyLacksClosingParagraph,
} from "./stories";
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
 * What each id the view minted stood in for: no id, or the id an earlier
 * paragraph of the package already holds.
 */
type MintedParagraphIds = ReadonlyMap<string, string | undefined>;

/**
 * `document` with an id of its own on every paragraph, which resolution
 * addresses paragraphs by, and the ids it minted: the view gives each
 * paragraph its own id back. A host model can repeat an id across stories
 * (a header copied from the body), where an operation refuses the repeat.
 */
const withParagraphIds = (
  document: Document,
): { document: Document; minted: MintedParagraphIds } => {
  const taken = new Set(packageParagraphIds(document.package));
  const minted = new Map<string, string | undefined>();
  const held = new Set<string>();
  let next = 0x10000000;
  const fresh = (original: string | undefined): string => {
    let id = next.toString(16).toUpperCase();
    while (taken.has(id)) {
      next += 1;
      id = next.toString(16).toUpperCase();
    }
    next += 1;
    minted.set(id, original);
    return id;
  };
  const pkg = mapRecords(document.package, (record) => {
    if (record["type"] !== "paragraph") return record;
    const id = record["paraId"];
    if (typeof id !== "string") return { ...record, paraId: fresh(undefined) };
    if (!held.has(id)) {
      held.add(id);
      return record;
    }
    return { ...record, paraId: fresh(id) };
  }) as Document["package"];
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
  let next = maxOf(ids) + 1;
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

/** The closing paragraphs the view gave stories, by id. */
type ClosingParagraphs = ReadonlyMap<string, OpStory>;

/**
 * `blocks` with the terminal-table review carrier they end with as the
 * paragraph it has become, where its own mark is an insertion: words were put
 * there, and rejecting them takes the paragraph. Otherwise `undefined`.
 */
const withInsertedCarrierPlain = (
  blocks: readonly BlockContent[],
): BlockContent[] | undefined => {
  const last = blocks.at(-1);
  if (
    last?.type !== "paragraph" ||
    last.reviewCarrier !== REVIEW_CARRIERS.TERMINAL_TABLE ||
    last.pPrMark?.kind !== "ins"
  )
    return undefined;
  const { reviewCarrier: _carrier, ...plain } = last;
  return [...blocks.slice(0, -1), plain];
};

/**
 * `document` with the paragraph a story ends with given to every story in
 * `stories` that lacks it, as a session reads such a story: resolution keeps
 * a container's final paragraph, so a story that ends in a table, or in a
 * carrier that has become an inserted paragraph, resolves with one there.
 */
const withClosingParagraphs = (
  document: Document,
  stories: readonly OpStory[],
): { document: Document; closing: ClosingParagraphs } => {
  const closing = new Map<string, OpStory>();
  let taken: Set<string> | undefined;
  let next = 0x20000000;
  let current = document;
  for (const story of stories) {
    const body = findStoryBody(current, story);
    if (body === undefined) continue;
    const carried = withInsertedCarrierPlain(body.content);
    if (carried === undefined && !storyLacksClosingParagraph(current, story))
      continue;
    taken ??= new Set(packageParagraphIds(document.package));
    let paraId = next.toString(16).toUpperCase();
    while (taken.has(paraId)) {
      next += 1;
      paraId = next.toString(16).toUpperCase();
    }
    next += 1;
    closing.set(paraId, story);
    current = replaceStoryBody({
      document: current,
      story,
      body: withBodyContent(body, [
        ...(carried ?? body.content),
        { type: "paragraph", paraId, content: [] },
      ]),
    });
  }
  return { document: current, closing };
};

const isBareParagraph = (block: BlockContent | undefined): boolean =>
  block?.type === "paragraph" &&
  block.content.length === 0 &&
  Object.keys(block).every(
    (key) => key === "type" || key === "paraId" || key === "content",
  );

/**
 * `document` with each closing paragraph the view gave a story taken out
 * again where resolution left it as given, last and bare, after other
 * blocks: the story keeps the ending it was read with. One that took
 * content, or is all its story holds, stays.
 */
const withoutBareClosingParagraphs = (
  document: Document,
  closing: ClosingParagraphs,
): Document => {
  let current = document;
  for (const [paraId, story] of closing) {
    const body = findStoryBody(current, story);
    if (body === undefined || body.content.length < 2) continue;
    const last = body.content.at(-1);
    if (last?.type !== "paragraph" || last.paraId !== paraId) continue;
    if (!isBareParagraph(last)) continue;
    current = replaceStoryBody({
      document: current,
      story,
      body: withBodyContent(body, body.content.slice(0, -1)),
    });
  }
  return current;
};

/**
 * `document` with the terminal-table review carrier `story` ends with
 * resolved by `decision`, once the story's changes are: the carrier stands in
 * for the paragraph a story needs after a table shown deleted. Accepted, or
 * holding content, it stays as a plain paragraph; rejected and empty, it goes
 * and the story ends in its table again, unless it is all the story holds.
 */
const withTerminalTableCarrierResolved = ({
  document,
  story,
  decision,
}: {
  document: Document;
  story: OpStory;
  decision: RevisionDecision;
}): Document => {
  const body = findStoryBody(document, story);
  const carrier = body?.content.at(-1);
  if (
    body === undefined ||
    carrier?.type !== "paragraph" ||
    carrier.reviewCarrier !== REVIEW_CARRIERS.TERMINAL_TABLE
  )
    return document;
  const { reviewCarrier: _carrier, ...plain } = carrier;
  const kept =
    decision === "accept" ||
    carrier.content.length > 0 ||
    body.content.length === 1;
  return replaceStoryBody({
    document,
    story,
    body: withBodyContent(
      body,
      kept ? [...body.content.slice(0, -1), plain] : body.content.slice(0, -1),
    ),
  });
};

const withoutParagraphIds = (
  document: Document,
  minted: MintedParagraphIds,
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
    const original = minted.get(id);
    if (original !== undefined) return { ...record, paraId: original };
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
  const revised = stories.filter(
    (story) => storyRevisionIds(document, story).length > 0,
  );
  if (revised.length === 0) return { document, unresolved: [] };
  const identified = withParagraphIds(withDistinctRevisionIds(document));
  const closed = withClosingParagraphs(identified.document, revised);
  let current = closed.document;
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
    else
      current = withTerminalTableCarrierResolved({
        document: resolved.value.document,
        story,
        decision,
      });
  }
  const minted = new Map(identified.minted);
  for (const paraId of closed.closing.keys()) minted.set(paraId, undefined);
  let resolvedDocument = withoutParagraphIds(
    withoutBareClosingParagraphs(current, closed.closing),
    minted,
  );
  // A refused story keeps the blocks it was read with, revision ids included.
  for (const story of unresolved) {
    const body = findStoryBody(document, story);
    if (body !== undefined)
      resolvedDocument = replaceStoryBody({
        document: resolvedDocument,
        story,
        body,
      });
  }
  return { document: resolvedDocument, unresolved };
};
