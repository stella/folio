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
 * inverses too. An inverse states what it expects to find (the slice it
 * removes, the values a patch replaced, the fields of a paragraph it merges
 * away) and is refused as stale when that has changed.
 *
 * Wire format. An operation is plain data and survives `JSON.stringify` and
 * `JSON.parse` unchanged; it is journaled inside a {@link DocumentOpEnvelope}
 * that names this schema. Operations embed model records (`Paragraph`,
 * `ParagraphContent`, property sets) as they stand in schema version 1, so
 * those record shapes are part of the wire format too: changing one changes
 * what a journaled operation means, and needs a new schema version and a
 * migration of the journaled operations.
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
 * nothing, empty containers) can sit at one offset. `zeroWidthBefore` says how
 * many of them, in document order, come before the point; each operation
 * states what it assumes when the field is absent. It is a second coordinate
 * of the position: a step that rebases an operation over one that inserted or
 * removed zero-width children at the same offset shifts it as it shifts
 * `offset` over inserted or removed units.
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

/**
 * Revision and content-control ids an operation gives the records it creates
 * by cutting an identified record in two (see `identity.ts`), each space's
 * taken in document order. The half holding the start keeps the record's id.
 */
export type NewIds = {
  /** For tracked changes and tracked property changes (`w:id`). */
  revision?: readonly number[];
  /** For content controls (`w:sdtPr/w:id`). */
  control?: readonly number[];
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
 * becomes a run of its own and that run is split around it; `newIds` names
 * the ids of the second half when the run carries some. The run decides where
 * the text goes, so `at.zeroWidthBefore` is not read.
 */
export type InsertTextOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_TEXT;
  at: TextPosition;
  /** Characters only: tabs, breaks and other inline atoms are not text. */
  text: string;
  runProps: InsertedRunProps;
  newIds?: NewIds;
};

/**
 * Insert paragraph content at a position: runs, fields, markers, anything a
 * paragraph holds.
 *
 * The content is cut in at the point, splitting the runs and containers the
 * point falls inside; the slice's open ends then merge with the halves. A
 * slice that would merge two records that were separate is refused.
 * Absent `zeroWidthBefore` puts the point before the first zero-width child
 * at the offset that opens a range, after the others.
 */
export type InsertContentOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_CONTENT;
  at: TextPosition;
  slice: InlineSlice;
  newIds?: NewIds;
};

/**
 * Remove what lies between two positions of one paragraph: its units and the
 * zero-width children strictly inside. Absent `zeroWidthBefore` keeps the
 * zero-width children at both ends. A run or container left with nothing is
 * removed; one the range passes through keeps its two ends as one record.
 *
 * `join` then merges the records meeting where the range was, that many
 * levels below the ones running across it, as {@link JoinInlineOp} does: the
 * inverse of an insertion that cut records is one deletion.
 *
 * `expected` is the slice the deletion must remove, compared without the
 * fields a relayout recomputes. It is how an inverse refuses to apply to
 * content that has changed since.
 */
export type DeleteRangeOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_RANGE;
  from: TextPosition;
  to: TextPosition;
  join?: number;
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
  newIds?: NewIds;
};

/**
 * Merge the records meeting at a position, `depth` levels below the ones that
 * already run across it. Merged records must be alike in everything but
 * their content and ids; the first's ids stand for both. Absent
 * `zeroWidthBefore` is `0`.
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
 *
 * `expected` states the values the patched keys must have on every run in the
 * range (`null` for absent) and refuses the patch as stale otherwise.
 * `joinStart` and `joinEnd` merge the runs meeting at the range ends that many
 * levels once the patch is applied, as {@link JoinInlineOp} does; the inverse
 * of a patch that cut runs uses them.
 */
export type SetRunPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_RUN_PROPS;
  from: TextPosition;
  to: TextPosition;
  patch: RunPropsPatch;
  whenEmpty?: EmptyPropertySet;
  expected?: RunPropsPatch;
  joinStart?: number;
  joinEnd?: number;
  newIds?: NewIds;
};

/**
 * Patch a paragraph's own property set. `expected` states the values the
 * patched keys must have (`null` for absent) and refuses the patch as stale
 * otherwise.
 */
export type SetParagraphPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS;
  story: OpStory;
  blockId: string;
  patch: ParagraphPropsPatch;
  whenEmpty?: EmptyPropertySet;
  expected?: ParagraphPropsPatch;
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
 *
 * A split inside a tracked change, a content control or a run with a tracked
 * property change continues it in the new paragraph as a record of its own:
 * `newIds` names its ids, outermost record first.
 */
export type SplitBlockOp = {
  type: typeof DOCUMENT_OP_TYPES.SPLIT_BLOCK;
  at: TextPosition;
  /** `w14:paraId` of the new paragraph: eight hex digits, unused in the document. */
  newBlockId: string;
  newParagraph?: SplitParagraphFields;
  firstMark?: ParagraphMarkChange;
  newIds?: NewIds;
};

/**
 * Join a paragraph with the paragraph that directly follows it in the same
 * container. The first keeps its identity and properties; the second's content
 * and paragraph mark (section break, tracked mark change) join it, and the
 * second's id is retired. `depth` merges that many levels of the records
 * meeting at the join, as {@link JoinInlineOp} does.
 *
 * `expectedSecond` states the second paragraph's own fields, which the join
 * discards, and refuses the join as stale when they differ (fields a relayout
 * recomputes are not compared).
 */
export type JoinBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.JOIN_BLOCKS;
  story: OpStory;
  blockId: string;
  nextBlockId: string;
  depth?: number;
  expectedSecond?: SplitParagraphFields;
};

/**
 * Replace a run of adjacent paragraphs with other paragraphs.
 *
 * `expected` is the paragraphs as they stand: they are found by `paraId` and
 * compared structurally (without the fields a relayout recomputes), so a
 * replacement authored against another state is refused rather than applied
 * over it. No other operation's inverse is a replacement.
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

/**
 * An operation as it is journaled and sent: the schema that reads it, and the
 * operation. A reader refuses an envelope of a schema it does not know rather
 * than guess what its fields mean.
 */
export type DocumentOpEnvelope = {
  schema: typeof DOCUMENT_OP_SCHEMA_VERSION;
  op: DocumentOp;
};

/** An operation in the envelope it is journaled and sent in. */
export const toOpEnvelope = (op: DocumentOp): DocumentOpEnvelope => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  op,
});

/** The blocks an operation changed, by id. */
export type TouchedBlocks = {
  modified: readonly string[];
  inserted: readonly string[];
  removed: readonly string[];
};
