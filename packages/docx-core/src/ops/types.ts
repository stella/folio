/**
 * Document operations, schema version 1: text and formatting edits on the
 * main story.
 *
 * An operation names everything it needs. Positions are `(story, blockId,
 * offset)`, where `blockId` is the paragraph's `w14:paraId` and `offset` counts
 * the paragraph's logical units (see `offsets.ts`). Ids a new record takes are
 * carried in the operation (`splitBlock.newBlockId`), so applying one never
 * mints an identifier, reads a clock or draws a random number: equal inputs
 * give equal outputs wherever the operation runs.
 *
 * The values are plain data. An operation survives `JSON.stringify` and
 * `JSON.parse` unchanged, which is what lets the same operation be journaled,
 * sent and replayed.
 */

import type { Paragraph, ParagraphFormatting, TextFormatting } from "../model/document";

/** The operation schema this module reads and writes. */
export const DOCUMENT_OP_SCHEMA_VERSION = 1;

/**
 * The stories an operation can address. Headers, footers, notes and comment
 * bodies are stories too; version 1 addresses the main story only.
 */
export const OP_STORIES = Object.freeze({ MAIN: "main" } as const);

/** One of {@link OP_STORIES}. */
export type OpStory = (typeof OP_STORIES)[keyof typeof OP_STORIES];

/** A point in one paragraph's logical offset space. */
export type TextPosition = {
  story: OpStory;
  /** The paragraph's `w14:paraId`. */
  blockId: string;
  /** Logical units from the start of the paragraph, `0` to its length. */
  offset: number;
};

/**
 * Per-key change to a property set: a value sets the key, `null` clears it,
 * and a key the patch leaves out is untouched.
 */
export type FormattingPatch<Formatting> = {
  readonly [Key in keyof Formatting]?: Exclude<Formatting[Key], undefined> | null;
};

/** A change to a run property set (`w:rPr`). */
export type RunPropsPatch = FormattingPatch<TextFormatting>;

/** A change to a paragraph property set (`w:pPr`). */
export type ParagraphPropsPatch = FormattingPatch<ParagraphFormatting>;

/**
 * Inserted text takes the properties of the run it joins: the run holding the
 * unit before the position, or the one after it at the start of a paragraph.
 */
export const INHERIT_RUN_PROPS = "inherit";

/** The run properties inserted text carries. */
export type InsertedRunProps = typeof INHERIT_RUN_PROPS | TextFormatting;

/** The operation kinds of schema version 1. */
export const DOCUMENT_OP_TYPES = Object.freeze({
  INSERT_TEXT: "insertText",
  DELETE_RANGE: "deleteRange",
  SET_RUN_PROPS: "setRunProps",
  SET_PARAGRAPH_PROPS: "setParagraphProps",
  SPLIT_BLOCK: "splitBlock",
  JOIN_BLOCKS: "joinBlocks",
  REPLACE_BLOCKS: "replaceBlocks",
} as const);

/** One of {@link DOCUMENT_OP_TYPES}. */
export type DocumentOpType = (typeof DOCUMENT_OP_TYPES)[keyof typeof DOCUMENT_OP_TYPES];

/**
 * Insert text at a position.
 *
 * `runProps` is either {@link INHERIT_RUN_PROPS} or the exact property set the
 * text carries. When the set differs from the run at the position, the text
 * becomes a run of its own and that run is split around it.
 */
export type InsertTextOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_TEXT;
  at: TextPosition;
  /** Characters only: tabs, breaks and other inline atoms are not text. */
  text: string;
  runProps: InsertedRunProps;
};

/**
 * Remove the units between two positions of one paragraph.
 *
 * Zero-width markers in the range (bookmark, comment and move boundaries, and
 * captured markup that shows nothing) stay: an anchor outlives the text it
 * covered. A run or inline container the deletion empties goes with it.
 */
export type DeleteRangeOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_RANGE;
  from: TextPosition;
  to: TextPosition;
};

/**
 * Patch the run properties of every run unit between two positions of one
 * paragraph, splitting runs at the range ends. A run the patch does not change
 * is kept whole.
 */
export type SetRunPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_RUN_PROPS;
  from: TextPosition;
  to: TextPosition;
  patch: RunPropsPatch;
};

/** Patch a paragraph's own property set. */
export type SetParagraphPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS;
  story: OpStory;
  blockId: string;
  patch: ParagraphPropsPatch;
};

/**
 * Split a paragraph in two at a position.
 *
 * The half before the position keeps the paragraph's identity: its `paraId`,
 * `textId`, unmodelled attributes and tracked property change. The half after
 * it is a new paragraph named `newBlockId`; it holds the paragraph mark, so the
 * section break and a tracked mark change move with it. It takes `newProps`,
 * or a copy of the paragraph's properties when the operation states none.
 */
export type SplitBlockOp = {
  type: typeof DOCUMENT_OP_TYPES.SPLIT_BLOCK;
  at: TextPosition;
  /** `w14:paraId` of the new paragraph: eight hex digits, unused in the story. */
  newBlockId: string;
  newProps?: ParagraphFormatting;
};

/**
 * Join a paragraph with the paragraph that directly follows it in the same
 * container. The first keeps its identity and properties; the second's content
 * and paragraph mark (section break, tracked mark change) join it, and the
 * second's id is retired.
 */
export type JoinBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.JOIN_BLOCKS;
  story: OpStory;
  blockId: string;
  nextBlockId: string;
};

/**
 * Replace a run of adjacent paragraphs with other paragraphs.
 *
 * `expected` is the paragraphs as they stand: they are found by `paraId` and
 * compared structurally, so a replacement authored against another state is
 * refused rather than applied over it. This is the exact inverse every other
 * operation records: `expected` is what the operation produced and `blocks`
 * is the original records.
 */
export type ReplaceBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.REPLACE_BLOCKS;
  story: OpStory;
  expected: readonly Paragraph[];
  blocks: readonly Paragraph[];
};

/** A schema-version-1 document operation. */
export type DocumentOp =
  | InsertTextOp
  | DeleteRangeOp
  | SetRunPropsOp
  | SetParagraphPropsOp
  | SplitBlockOp
  | JoinBlocksOp
  | ReplaceBlocksOp;

/** The blocks an operation changed, by id. */
export type TouchedBlocks = {
  modified: readonly string[];
  inserted: readonly string[];
  removed: readonly string[];
};
