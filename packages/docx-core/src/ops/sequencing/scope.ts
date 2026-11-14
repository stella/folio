import { panic, Result } from "better-result";

import type { Document } from "../../model/document";
import { storyBody, storyParagraphs } from "../blocks";
import { structurallyEqual } from "../equality";
import { idKey } from "../ids";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
  type TouchedBlocks,
} from "../types";
import {
  anchorPoints,
  isAnchoredComment,
  isCommentOp,
  transformCommentPair,
} from "./commentTransform";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  type BatchRejectionReason,
  EXCLUSIVE_PARTS,
  type ExclusivePart,
  type SequencedOpEffect,
  TABLE_EXCLUSIVE_OP_TYPES,
} from "./envelope";
import { isInlineInsertion } from "./points";

type PartAccess = {
  part: ExclusivePart;
  /** One entry of the part; absent names the whole part. */
  key?: string;
  write: boolean;
};

/** What an operation addresses, as far as the operation itself says. */
type OpFootprint = {
  /** Rebased by the pairwise position transform rather than by footprint. */
  positional: boolean;
  /** Paragraph ids the operation addresses, reads or creates. */
  blocks: readonly string[];
  parts: readonly PartAccess[];
};

/**
 * Whether the operation changes the structure of the table holding its
 * paragraph: rows, columns, cells, grid or the table's own properties.
 */
export const restructuresTable = (op: DocumentOp): boolean =>
  Object.hasOwn(TABLE_EXCLUSIVE_OP_TYPES, op.type);

type PayloadFacts = { paraIds: string[]; numIds: number[]; sections: boolean };

// Paragraph ids, numbering references and section properties an embedded
// payload carries, in any depth of nested content.
const payloadFacts = (value: unknown): PayloadFacts => {
  const facts: PayloadFacts = { paraIds: [], numIds: [], sections: false };
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const child of node) visit(child);
      return;
    }
    if (typeof node !== "object" || node === null) return;
    for (const [key, child] of Object.entries(node)) {
      if (key === "paraId" && typeof child === "string")
        facts.paraIds.push(child);
      if (
        key === "numPr" &&
        typeof child === "object" &&
        child !== null &&
        "numId" in child &&
        typeof child.numId === "number"
      )
        facts.numIds.push(child.numId);
      if (key === "sectionProperties" && child !== undefined)
        facts.sections = true;
      visit(child);
    }
  };
  visit(value);
  return facts;
};

const write = (part: ExclusivePart, key?: string): PartAccess =>
  key === undefined ? { part, write: true } : { part, key, write: true };
const read = (part: ExclusivePart, key: string): PartAccess => ({
  part,
  key,
  write: false,
});

const sections = (changes: boolean): PartAccess[] =>
  changes ? [write(EXCLUSIVE_PARTS.SECTIONS)] : [];

const comment = (id: number): string => String(id);

/** The footprint of every operation kind. */
const footprintOf = (op: DocumentOp): OpFootprint => {
  const facts = payloadFacts(op);
  const numbering = facts.numIds.map((id) =>
    read(EXCLUSIVE_PARTS.NUMBERING, `num:${String(id)}`),
  );
  const scoped = (
    blocks: readonly (string | undefined)[],
    parts: readonly PartAccess[] = [],
  ): OpFootprint => ({
    positional: false,
    blocks: [
      ...blocks.filter((id): id is string => id !== undefined),
      ...facts.paraIds,
    ],
    parts: [...parts, ...numbering, ...sections(facts.sections)],
  });
  const positional = (
    blocks: readonly string[],
    parts: readonly PartAccess[] = [],
  ): OpFootprint => ({
    positional: true,
    blocks: [...blocks, ...facts.paraIds],
    parts: [...parts, ...numbering, ...sections(facts.sections)],
  });
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
      return positional([op.at.blockId]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return positional([op.from.blockId, op.to.blockId]);
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return positional([op.blockId]);
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return positional(
        [op.at.blockId, op.newBlockId],
        sections(
          op.firstSectionProperties !== undefined ||
            op.sectionView !== undefined,
        ),
      );
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return positional(
        [op.blockId, op.nextBlockId],
        sections(
          op.sectionBoundary !== undefined || op.sectionView !== undefined,
        ),
      );
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return positional([op.at.blockId]);
    // Decisions address revisions by id: no offset of theirs moves.
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
      return positional(
        [],
        op.revisionIds.map((id) => read(EXCLUSIVE_PARTS.REVISIONS, String(id))),
      );
    // Inline content moves offsets like text; a slice that merges with the
    // records at its point depends on them and is block-scoped.
    case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
      const blocks = [op.at.blockId];
      return isInlineInsertion(op) ? positional(blocks) : scoped(blocks);
    }
    // Cutting or merging inline records moves no offset.
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return positional([op.at.blockId]);
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
      return scoped(op.blockIds);
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return scoped(
        [],
        sections(
          op.sectionBoundaries !== undefined || op.sectionView !== undefined,
        ),
      );
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE:
      return scoped([
        op.blockId,
        ...("newBlockIds" in op ? op.newBlockIds : []),
      ]);
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
      return scoped([op.at.blockId, op.terminal?.beforeBlockId]);
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE: {
      const whole =
        op.expected !== undefined ||
        op.restore !== undefined ||
        op.undefinedFields?.some(
          ([field]) => field === "expected" || field === "restore",
        ) === true;
      return scoped(
        [],
        whole
          ? [write(EXCLUSIVE_PARTS.NUMBERING)]
          : [
              write(EXCLUSIVE_PARTS.NUMBERING, `num:${String(op.num.numId)}`),
              write(
                EXCLUSIVE_PARTS.NUMBERING,
                `abstract:${String(op.abstractNum?.abstractNumId ?? op.num.abstractNumId)}`,
              ),
            ],
      );
    }
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return scoped(
        [op.endpoint.type === "paragraph" ? op.endpoint.blockId : undefined],
        [write(EXCLUSIVE_PARTS.SECTIONS)],
      );
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
      return scoped([], [write(EXCLUSIVE_PARTS.SECTIONS)]);
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
      return scoped(
        [],
        [
          ...(structurallyEqual(op.expected.styles, op.resources.styles)
            ? []
            : [write(EXCLUSIVE_PARTS.STYLES)]),
          ...(structurallyEqual(op.expected.numbering, op.resources.numbering)
            ? []
            : [write(EXCLUSIVE_PARTS.NUMBERING)]),
          ...op.relationships.entries.map(({ key }) =>
            write(EXCLUSIVE_PARTS.RELATIONSHIPS, key),
          ),
          ...op.media.entries.map(({ key }) =>
            write(EXCLUSIVE_PARTS.MEDIA, key),
          ),
          ...(op.relationships.expected === op.relationships.next
            ? []
            : [write(EXCLUSIVE_PARTS.RELATIONSHIPS)]),
          ...(op.media.expected === op.media.next
            ? []
            : [write(EXCLUSIVE_PARTS.MEDIA)]),
        ],
      );
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
      return scoped(
        [],
        [
          write(EXCLUSIVE_PARTS.SECTIONS),
          write(EXCLUSIVE_PARTS.HEADERS_FOOTERS, op.story.rId),
          write(EXCLUSIVE_PARTS.RELATIONSHIPS, op.story.rId),
        ],
      );
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
      return scoped(
        [],
        [
          write(EXCLUSIVE_PARTS.SECTIONS),
          write(EXCLUSIVE_PARTS.HEADERS_FOOTERS),
          write(EXCLUSIVE_PARTS.RELATIONSHIPS),
          write(EXCLUSIVE_PARTS.NOTES),
          write(EXCLUSIVE_PARTS.SETTINGS),
          write(EXCLUSIVE_PARTS.COMMENTS),
        ],
      );
    case DOCUMENT_OP_TYPES.ADD_NOTE:
      return scoped(
        [op.at.blockId],
        [write(EXCLUSIVE_PARTS.NOTES, `${op.note.type}:${String(op.note.id)}`)],
      );
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
      return scoped(
        [op.at.blockId],
        [
          write(
            EXCLUSIVE_PARTS.NOTES,
            `${op.story.kind}:${String(op.story.id)}`,
          ),
        ],
      );
    case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
      return scoped(
        [],
        op.infos.map(({ id }) => write(EXCLUSIVE_PARTS.REVISIONS, String(id))),
      );
    // Comment operations may add or remove the comment part relationships;
    // a range or point anchor moves like text.
    case DOCUMENT_OP_TYPES.CREATE_COMMENT: {
      const parts = [
        write(EXCLUSIVE_PARTS.RELATIONSHIPS),
        write(EXCLUSIVE_PARTS.COMMENTS, comment(op.comment.id)),
      ];
      if (isAnchoredComment(op))
        return positional(
          anchorPoints(op).map(({ blockId }) => blockId),
          parts,
        );
      switch (op.anchor.kind) {
        case "reply":
          return scoped(
            [],
            [
              ...parts,
              read(EXCLUSIVE_PARTS.COMMENTS, comment(op.anchor.parentId)),
            ],
          );
        case "revision":
          return scoped(
            [],
            [
              ...parts,
              write(EXCLUSIVE_PARTS.REVISIONS, String(op.anchor.revisionId)),
            ],
          );
        default:
          return panic("Unknown comment anchor kind.");
      }
    }
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
      return scoped([], [write(EXCLUSIVE_PARTS.COMMENTS, comment(op.id))]);
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
      return scoped(
        [],
        [
          write(EXCLUSIVE_PARTS.RELATIONSHIPS),
          write(EXCLUSIVE_PARTS.COMMENTS, comment(op.id)),
        ],
      );
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      return scoped(
        [...op.expected.anchors, ...op.state.anchors].map(
          ({ blockId }) => blockId,
        ),
        [
          write(EXCLUSIVE_PARTS.RELATIONSHIPS),
          ...op.ids.map((id) => write(EXCLUSIVE_PARTS.COMMENTS, comment(id))),
        ],
      );
    default: {
      const exhaustive: never = op;
      void exhaustive;
      return panic("Unknown document operation kind.");
    }
  }
};

/** Whether the operation is sequenced with a block footprint, not a position transform. */
export const isBlockScoped = (op: DocumentOp): boolean =>
  !footprintOf(op).positional;

const partsConflict = (
  left: readonly PartAccess[],
  right: readonly PartAccess[],
): PartAccess | undefined =>
  left.find((mine) =>
    right.some(
      (theirs) =>
        mine.part === theirs.part &&
        (mine.write || theirs.write) &&
        (mine.key === undefined ||
          theirs.key === undefined ||
          mine.key === theirs.key),
    ),
  );

type ScopedPair = {
  op: DocumentOp;
  over: DocumentOp;
  effect: SequencedOpEffect | undefined;
  order: "before" | "after";
};

export const TOUCHED_BLOCKS_EFFECT = "touchedBlocks";

const keys = (ids: readonly string[]): Set<string> => new Set(ids.map(idKey));

/**
 * Decide a pair by footprint. Undefined leaves the pair to the position
 * transform: both operations rebase by offset and share no exclusive part.
 * A block-scoped operation passes a concurrent one unchanged when it
 * addresses none of the paragraphs that one addressed or changed; it is
 * refused when it does, and a table structure edit is refused when the
 * concurrent one restructured the same table.
 */
export const transformByFootprint = ({
  op,
  over,
  effect,
  order,
}: ScopedPair): Result<readonly DocumentOp[], BatchRejection> | undefined => {
  // Two comment operations decide by the comments they use and their anchors.
  if (isCommentOp(op) && isCommentOp(over))
    return transformCommentPair({ op, over, effect, order });
  const mine = footprintOf(op);
  const theirs = footprintOf(over);
  const refuse = (
    reason: BatchRejectionReason,
    message: string,
    part?: ExclusivePart,
  ) =>
    Result.err(
      new BatchRejection({
        reason,
        message,
        opType: op.type,
        overType: over.type,
        ...(part === undefined ? {} : { part }),
      }),
    );
  const shared = partsConflict(mine.parts, theirs.parts);
  if (shared !== undefined)
    return refuse(
      BATCH_REJECTION_REASONS.PART_CONFLICT,
      `A concurrent operation changed the ${shared.part} this operation uses.`,
      shared.part,
    );
  if (mine.positional && theirs.positional) return undefined;
  const touched = effect?.type === TOUCHED_BLOCKS_EFFECT ? effect : undefined;
  const changed = keys([...theirs.blocks, ...(touched?.blockIds ?? [])]);
  const addressed = mine.blocks.map(idKey);
  if (
    restructuresTable(op) &&
    addressed.some((id) => keys(touched?.tableBlockIds ?? []).has(id))
  )
    return refuse(
      BATCH_REJECTION_REASONS.TABLE_REQUIRES_EXCLUSIVE_EDIT,
      "A concurrent operation restructured the same table.",
    );
  if (addressed.some((id) => changed.has(id)))
    return refuse(
      BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
      "A concurrent operation changed a paragraph this operation addresses.",
    );
  return Result.ok([op]);
};

const operationStory = (op: DocumentOp): OpStory => {
  if ("at" in op && typeof op.at === "object" && "story" in op.at)
    return op.at.story;
  if ("from" in op) return op.from.story;
  if ("story" in op && typeof op.story === "string") return op.story;
  return OP_STORIES.MAIN;
};

const paragraphsOf = (document: Document, story: OpStory) =>
  new Map(
    storyParagraphs(storyBody(document, story)).flatMap(({ paragraph }) =>
      paragraph.paraId === undefined
        ? []
        : [[idKey(paragraph.paraId), paragraph] as const],
    ),
  );

// Every paragraph of the innermost table holding the operation's paragraph.
const tableParagraphIds = (
  document: Document,
  op: DocumentOp,
): readonly string[] => {
  if (!("blockId" in op)) return [];
  const locations = storyParagraphs(storyBody(document, operationStory(op)));
  const target = locations.find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.blockId),
  );
  const step = target?.list.findLastIndex(({ kind }) => kind === "tableCell");
  if (target === undefined || step === undefined || step < 0) return [];
  const table = JSON.stringify([
    ...target.list.slice(0, step),
    target.list[step]?.block,
  ]);
  return locations.flatMap(({ list, paragraph }) => {
    const cell = list[step];
    return paragraph.paraId !== undefined &&
      cell?.kind === "tableCell" &&
      JSON.stringify([...list.slice(0, step), cell.block]) === table
      ? [paragraph.paraId]
      : [];
  });
};

// The comments whose records an operation added, removed or replaced, and
// whether a package relationship went with them.
const changedComments = (
  before: Document,
  after: Document,
): { ids: number[]; relationshipRemoved: boolean } => {
  const previous = new Map(
    (before.package.document.comments ?? []).map((record) => [
      record.id,
      record,
    ]),
  );
  const next = new Map(
    (after.package.document.comments ?? []).map((record) => [
      record.id,
      record,
    ]),
  );
  const ids = new Set<number>();
  for (const [id, record] of previous) if (next.get(id) !== record) ids.add(id);
  for (const id of next.keys()) if (!previous.has(id)) ids.add(id);
  return {
    ids: [...ids].toSorted((left, right) => left - right),
    relationshipRemoved: [...(before.package.relationships?.keys() ?? [])].some(
      (key) => after.package.relationships?.has(key) !== true,
    ),
  };
};

/**
 * The sequenced effect of a block-scoped operation: every paragraph it
 * changed, inserted or removed, and for a table structure edit every
 * paragraph of the table before and after it.
 */
export const blockScopedEffect = (
  before: Document,
  after: Document,
  op: DocumentOp,
  touched: TouchedBlocks,
): SequencedOpEffect => {
  const story = operationStory(op);
  const previous = paragraphsOf(before, story);
  const next = paragraphsOf(after, story);
  const changed = new Set<string>();
  for (const [key, paragraph] of previous)
    if (next.get(key) !== paragraph) changed.add(key);
  for (const key of next.keys()) if (!previous.has(key)) changed.add(key);
  const ids = new Map<string, string>();
  for (const id of [
    ...touched.modified,
    ...touched.inserted,
    ...touched.removed,
  ])
    ids.set(idKey(id), id);
  for (const key of changed) {
    const paragraph = previous.get(key) ?? next.get(key);
    if (paragraph?.paraId !== undefined && !ids.has(key))
      ids.set(key, paragraph.paraId);
  }
  const blockIds = [...ids.values()].toSorted();
  if (isCommentOp(op))
    return {
      type: TOUCHED_BLOCKS_EFFECT,
      blockIds,
      comments: changedComments(before, after),
    };
  if (!restructuresTable(op)) return { type: TOUCHED_BLOCKS_EFFECT, blockIds };
  const table = [
    ...new Set([
      ...tableParagraphIds(before, op),
      ...tableParagraphIds(after, op),
    ]),
  ].toSorted();
  return { type: TOUCHED_BLOCKS_EFFECT, blockIds, tableBlockIds: table };
};
