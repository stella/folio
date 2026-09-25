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
 * Every operation's inverse is an operation of the same schema, addressed the
 * same way, so a step that rebases operations over concurrent edits rebases
 * inverses too. The values are plain data: an operation survives
 * `JSON.stringify` and `JSON.parse` unchanged.
 *
 * A run cut in two keeps every field on both halves, a tracked property
 * change (`w:rPrChange`) and its `w:id` included: the halves are one revision,
 * as the run was one. Operations mint no ids, so separating them is left to
 * whoever issues ids.
 */

import type {
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphMarkChange,
  TextFormatting,
} from "../model/document";

/** The operation schema this module reads and writes. */
export const DOCUMENT_OP_SCHEMA_VERSION = 1;

/**
 * The stories an operation can address. Headers, footers, notes and comment
 * bodies are stories too; version 1 addresses the main story only.
 */
export const OP_STORIES = Object.freeze({ MAIN: "main" } as const);

/** One of {@link OP_STORIES}. */
export type OpStory = (typeof OP_STORIES)[keyof typeof OP_STORIES];

/**
 * A point in one paragraph's logical offset space.
 *
 * Several zero-width children (range markers, captured markup that shows
 * nothing, empty runs) can sit at one offset. `zeroWidthBefore` says how many
 * of them, in document order, come before the point; each operation states
 * what it assumes when the field is absent.
 */
export type TextPosition = {
  story: OpStory;
  /** The paragraph's `w14:paraId`. */
  blockId: string;
  /** Logical units from the start of the paragraph, `0` to its length. */
  offset: number;
  zeroWidthBefore?: number;
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
 * What a property set a patch leaves without keys becomes: absent, or an
 * empty set. The two state the same thing; the choice is how an inverse
 * gives back a record in the spelling it had.
 */
export const EMPTY_PROPERTY_SETS = Object.freeze({ OMIT: "omit", KEEP: "keep" } as const);

/** One of {@link EMPTY_PROPERTY_SETS}. */
export type EmptyPropertySet = (typeof EMPTY_PROPERTY_SETS)[keyof typeof EMPTY_PROPERTY_SETS];

/**
 * Inserted text takes the properties of the run it joins: the run holding the
 * unit before the position, or the one after it at the start of a paragraph.
 */
export const INHERIT_RUN_PROPS = "inherit";

/** The run properties inserted text carries. */
export type InsertedRunProps = typeof INHERIT_RUN_PROPS | TextFormatting;

/**
 * Paragraph content with its cut ends marked.
 *
 * `openStart` counts the levels of the first item's left edge that continue
 * content before the insertion point: `0` puts the items in as they are, `1`
 * merges the first item with the one ending at the point, `2` also merges
 * their first and last children, and so on. `openEnd` is the same for the
 * last item and the content after the point. It is the shape a deletion
 * removes, and the shape that puts it back.
 */
export type InlineSlice = {
  content: readonly ParagraphContent[];
  openStart: number;
  openEnd: number;
};

/** The paragraph fields a split gives the new paragraph, besides its id, content and mark. */
export type SplitParagraphFields = Omit<
  Paragraph,
  "type" | "paraId" | "content" | "sectionProperties" | "pPrMark"
>;

/** The operation kinds of schema version 1. */
export const DOCUMENT_OP_TYPES = Object.freeze({
  INSERT_TEXT: "insertText",
  INSERT_CONTENT: "insertContent",
  DELETE_RANGE: "deleteRange",
  SPLIT_INLINE: "splitInline",
  JOIN_INLINE: "joinInline",
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
 * becomes a run of its own and that run is split around it. The run decides
 * where the text goes, so `at.zeroWidthBefore` is not read.
 */
export type InsertTextOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_TEXT;
  at: TextPosition;
  /** Characters only: tabs, breaks and other inline atoms are not text. */
  text: string;
  runProps: InsertedRunProps;
};

/**
 * Insert paragraph content at a position: runs, fields, markers, anything a
 * paragraph holds.
 *
 * The content is cut in at the point, splitting the runs and containers the
 * point falls inside; the slice's open ends then merge with the halves.
 * Absent `zeroWidthBefore` puts the point before the first zero-width child
 * at the offset that opens a range, after the others.
 */
export type InsertContentOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_CONTENT;
  at: TextPosition;
  slice: InlineSlice;
};

/**
 * Remove what lies between two positions of one paragraph: its units and the
 * zero-width children strictly inside. Absent `zeroWidthBefore` keeps the
 * zero-width children at both ends. A run or container left with nothing is
 * removed; one the range passes through keeps its two ends as one record.
 *
 * `expected` is the slice the deletion must remove. It is how an inverse
 * refuses to apply to content that has changed since.
 */
export type DeleteRangeOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_RANGE;
  from: TextPosition;
  to: TextPosition;
  expected?: InlineSlice;
};

/**
 * Cut the innermost `depth` records that run across a position in two,
 * leaving the ones around them whole. Absent `zeroWidthBefore` is `0`.
 */
export type SplitInlineOp = {
  type: typeof DOCUMENT_OP_TYPES.SPLIT_INLINE;
  at: TextPosition;
  depth: number;
};

/**
 * Merge the records meeting at a position, `depth` levels below the ones that
 * already run across it. Merged records must be alike in everything but
 * their content. Absent `zeroWidthBefore` is `0`.
 */
export type JoinInlineOp = {
  type: typeof DOCUMENT_OP_TYPES.JOIN_INLINE;
  at: TextPosition;
  depth: number;
};

/**
 * Patch the run properties of every run between two positions of one
 * paragraph, splitting runs at the range ends. A run the patch does not change
 * is kept whole. Absent `zeroWidthBefore` keeps zero-width children at both
 * ends out of the range.
 */
export type SetRunPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_RUN_PROPS;
  from: TextPosition;
  to: TextPosition;
  patch: RunPropsPatch;
  whenEmpty?: EmptyPropertySet;
};

/** Patch a paragraph's own property set. */
export type SetParagraphPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS;
  story: OpStory;
  blockId: string;
  patch: ParagraphPropsPatch;
  whenEmpty?: EmptyPropertySet;
};

/**
 * Split a paragraph in two at a position.
 *
 * The half before the position keeps the paragraph's identity: its `paraId`,
 * `textId`, unmodelled attributes and tracked property change. The half after
 * it is a new paragraph named `newBlockId`; it holds the paragraph mark, so the
 * section break and a tracked mark change move with it. It takes the fields in
 * `newParagraph`, or a copy of the paragraph's properties when the operation
 * states none. `firstMark` is the tracked change of the mark the split
 * creates. Absent `zeroWidthBefore` keeps zero-width children that open a
 * range with the new paragraph.
 */
export type SplitBlockOp = {
  type: typeof DOCUMENT_OP_TYPES.SPLIT_BLOCK;
  at: TextPosition;
  /** `w14:paraId` of the new paragraph: eight hex digits, unused in the document. */
  newBlockId: string;
  newParagraph?: SplitParagraphFields;
  firstMark?: ParagraphMarkChange;
};

/**
 * Join a paragraph with the paragraph that directly follows it in the same
 * container. The first keeps its identity and properties; the second's content
 * and paragraph mark (section break, tracked mark change) join it, and the
 * second's id is retired. `depth` merges that many levels of the records
 * meeting at the join, as {@link JoinInlineOp} does.
 */
export type JoinBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.JOIN_BLOCKS;
  story: OpStory;
  blockId: string;
  nextBlockId: string;
  depth?: number;
};

/**
 * Replace a run of adjacent paragraphs with other paragraphs.
 *
 * `expected` is the paragraphs as they stand: they are found by `paraId` and
 * compared structurally, so a replacement authored against another state is
 * refused rather than applied over it. No other operation's inverse is a
 * replacement.
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
  | InsertContentOp
  | DeleteRangeOp
  | SplitInlineOp
  | JoinInlineOp
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
