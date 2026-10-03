/**
 * Document operations, schema version 5: text, formatting and review edits on
 * independently addressed document stories, direct or tracked.
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
 * `ParagraphContent`, property sets) as they stand in this schema version, so
 * those record shapes are part of the wire format too: changing one changes
 * what a journaled operation means, and needs a new schema version and a
 * migration of the journaled operations.
 */

import type {
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphMarkChange,
  ParagraphPropertyChange,
  TextFormatting,
  TableRow,
  Table,
  BlockContent,
  HeaderFooterType,
  HeaderFooter,
  Footnote,
  Endnote,
  SectionProperties,
  DocumentBody,
  DocumentSettings,
  Section,
} from "../model/document";

/** JSON-safe section metadata; content is derived from the body's canonical blocks. */
export type SectionViewEntry = {
  properties: SectionProperties;
  headers?: readonly (readonly [HeaderFooterType, HeaderFooter])[];
  footers?: readonly (readonly [HeaderFooterType, HeaderFooter])[];
};

type SectionViewChange = {
  expected: readonly SectionViewEntry[];
  restore: readonly SectionViewEntry[];
};

/**
 * The operation schema this module reads and writes.
 *
 * Version 5 adds explicit section-boundary removal/restoration.
 * Version 5 adds header/footer/note story addresses and exact lifecycle operations.
 * Version 4 adds paragraph deletion through `deleteBlocks`, direct and tracked
 * whole-table operations with their exact structural inverse, terminal insertion,
 * and cell-ending paragraph marks on tracked row operations.
 * Version 3 added paragraph insertion through `insertBlocks` and the table-row
 * operations `insertRow`, `deleteRow` and `setTableRows`.
 * Version 2 added tracked changes: the `revision` stamp on the text, formatting
 * and paragraph operations, and the review operations `setParagraphReview`,
 * `replaceInline` and `resolveRevision`.
 */
export const DOCUMENT_OP_SCHEMA_VERSION = 5;

/**
 * The main story has a fixed address; other editable parts use their stable
 * relationship id or note id. Comment bodies are not yet editable stories.
 */
export const OP_STORIES = Object.freeze({ MAIN: "main" } as const);

/** The stable identity of an independently editable package story. */
export type HeaderFooterStory = { kind: "header" | "footer"; rId: string };
export type NoteStory = { kind: "footnote" | "endnote"; id: number };
export type OpStory = typeof OP_STORIES.MAIN | HeaderFooterStory | NoteStory;

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

/**
 * Who made a tracked change and when, and the revision id (`w:id`) of the
 * first record the change creates.
 *
 * An operation carrying a stamp records its edit as a tracked change instead
 * of making it: the stamp's id names the first record the edit creates and
 * `newIds.revision` names the others, in document order. The id must be unused
 * in the package. A new tracked wrapper that would sit next to one with the
 * same author, date and initials merges into it and creates no record, so an
 * author's burst of typing reads as one change.
 */
export type RevisionStamp = {
  id: number;
  author: string;
  /** `w:date`, as the author's client stated it. */
  date: string;
  initials?: string;
};

/** The paragraph fields a split gives the new paragraph, besides its id, content and mark. */
export type SplitParagraphFields = Omit<
  Paragraph,
  "type" | "paraId" | "content" | "sectionProperties" | "pPrMark"
>;

/** Explicit structural and review policies for editor intent compilation. */
export const SECTION_BOUNDARY_POLICIES = Object.freeze({
  REMOVE: "remove",
  REPLACE: "replace",
} as const);

/** The operation kinds of schema version 5. */
export const DOCUMENT_OP_TYPES = Object.freeze({
  CREATE_HEADER_FOOTER: "createHeaderFooter",
  REMOVE_HEADER_FOOTER: "removeHeaderFooter",
  ADD_NOTE: "addNote",
  REMOVE_NOTE: "removeNote",
  SET_SECTION_PROPS: "setSectionProps",
  RESTORE_STORY_PARTS: "restoreStoryParts",
  DELETE_BLOCKS: "deleteBlocks",
  INSERT_BLOCKS: "insertBlocks",
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
  SET_PARAGRAPH_REVIEW: "setParagraphReview",
  REPLACE_INLINE: "replaceInline",
  RESOLVE_REVISION: "resolveRevision",
  INSERT_TABLE: "insertTable",
  DELETE_TABLE: "deleteTable",
  SET_CONTAINER_BLOCKS: "setContainerBlocks",
  INSERT_ROW: "insertRow",
  DELETE_ROW: "deleteRow",
  SET_TABLE_ROWS: "setTableRows",
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
 *
 * With `revision`, the text is a tracked insertion: it is wrapped in an
 * `Insertion` carrying the stamp, cutting the records around it down to a
 * level that admits one. Text landing inside another stamp's insertion or
 * move splits that record around it rather than nesting inside it.
 */
export type InsertTextOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_TEXT;
  at: TextPosition;
  /** Characters only: tabs, breaks and other inline atoms are not text. */
  text: string;
  runProps: InsertedRunProps;
  newIds?: NewIds;
  revision?: RevisionStamp;
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
 *
 * With `revision`, the inserted content is a tracked insertion, as for
 * {@link InsertTextOp}. It is refused inside tracked-removed content, and
 * when it holds a comment boundary or reference, which no tracked change can
 * hold.
 */
export type InsertContentOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_CONTENT;
  at: TextPosition;
  slice: InlineSlice;
  newIds?: NewIds;
  revision?: RevisionStamp;
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
 *
 * With `revision`, the deletion is tracked and removes nothing. Content
 * already inside a deletion or moved away is left as it is; content inside an
 * insertion or a move gets a deletion nested inside that record; everything
 * else is wrapped in a `Deletion` carrying the stamp, an inline content
 * control the range covers whole going in with its content. A range holding
 * a comment boundary or reference is refused, since no tracked change can
 * hold one; `planTrackedDeletion` plans the deletions around them. A tracked
 * deletion joins nothing, so `join` must be absent.
 */
export type DeleteRangeOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_RANGE;
  from: TextPosition;
  to: TextPosition;
  join?: number;
  expected?: InlineSlice;
  newIds?: NewIds;
  revision?: RevisionStamp;
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
 *
 * With `revision`, every run the patch changes records a tracked property
 * change (`w:rPrChange`) whose previous formatting is the run's own property
 * set before the patch. A run already carrying one keeps it, and with it the
 * formatting it started from. A tracked patch joins nothing, so `joinStart`
 * and `joinEnd` must be absent.
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
  revision?: RevisionStamp;
};

/**
 * Patch a paragraph's own property set. `expected` states the values the
 * patched keys must have (`null` for absent) and refuses the patch as stale
 * otherwise.
 *
 * With `revision`, the paragraph records a tracked property change
 * (`w:pPrChange`) whose previous formatting is its own property set before
 * the patch, or folds into the one it already carries: its original baseline and
 * id stay, and its author/date become the latest stamp. A property change holds
 * paragraph properties only, not the paragraph mark's run properties, so a
 * tracked patch of {@link PARAGRAPH_MARK_FORMATTING_KEYS} is refused.
 */
export type SetParagraphPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS;
  story: OpStory;
  blockId: string;
  patch: ParagraphPropsPatch;
  whenEmpty?: EmptyPropertySet;
  expected?: ParagraphPropsPatch;
  revision?: RevisionStamp;
};

/**
 * Paragraph formatting that belongs to the paragraph mark's run properties
 * (`w:pPr/w:rPr`) rather than to the paragraph properties a `w:pPrChange`
 * records.
 */
export const PARAGRAPH_MARK_FORMATTING_KEYS = Object.freeze([
  "runProperties",
  "runInWithNext",
] as const satisfies readonly (keyof ParagraphFormatting)[]);

/** Which half of a split, or of a join, a rule applies to. */
export const SPLIT_HALVES = Object.freeze({ FIRST: "first", SECOND: "second" } as const);

/** One of {@link SPLIT_HALVES}. */
export type SplitHalf = (typeof SPLIT_HALVES)[keyof typeof SPLIT_HALVES];

/**
 * Split a paragraph in two at a position.
 *
 * One half keeps the paragraph's identity and own fields (`paraId`, `textId`,
 * unmodelled attributes, properties); the other is a new paragraph named
 * `newBlockId`, taking the fields in `newParagraph`, or a copy of the
 * paragraph's properties when the operation states none. `newHalf` says which
 * half is new: by default the second when the position ends the paragraph
 * (a paragraph added after it), else the first (the text before the position
 * moves to a paragraph of its own, and the paragraph continues with the rest).
 * The paragraph mark always ends the second half, so the section break, the
 * mark's tracked change and the pending property changes go with it; when the
 * second half is the new one, `newParagraph` states no property changes.
 * `firstMark` is the tracked change of the mark the split creates, which ends
 * the first half. Absent `zeroWidthBefore` keeps zero-width children that
 * open a range with the second half.
 *
 * A split inside a tracked change, a content control or a run with a tracked
 * property change continues it in the second half as a record of its own:
 * `newIds` names its ids, outermost record first.
 *
 * With `revision`, the mark the split creates is a tracked insertion
 * (`firstMark` must then be absent), and the new paragraph, when its paragraph
 * properties differ from the source's and it carries no property change,
 * records one from the source's. Rejecting the mark removes it, which leaves
 * the second half with the first's content before its own. A property change
 * does not record the mark's run properties, so a new second half, which ends
 * with the paragraph's own mark, must state the same ones as the paragraph.
 */
export type SplitBlockOp = {
  type: typeof DOCUMENT_OP_TYPES.SPLIT_BLOCK;
  at: TextPosition;
  /** `w14:paraId` of the new paragraph: eight hex digits, unused in the document. */
  newBlockId: string;
  newHalf?: SplitHalf;
  newParagraph?: SplitParagraphFields;
  firstMark?: ParagraphMarkChange;
  /** Exact inverse restoration of a removed section boundary on the first half. */
  firstSectionProperties?: SectionProperties;
  sectionView?: SectionViewChange;
  newIds?: NewIds;
  revision?: RevisionStamp;
};

/**
 * Join a paragraph with the paragraph that directly follows it in the same
 * container. The joined paragraph has the first's paragraph properties, or
 * the second's when the first holds no content, and the second's mark: its
 * run properties, section break, tracked change and the pending property
 * changes. It keeps the identity and own fields of the
 * `survivor` half, by default the second, and the other's id is retired.
 * An absent or zero `depth` merges the two plain runs at the seam when the
 * parser would merge them. A positive `depth` merges that many levels of the
 * records meeting at the join, as {@link JoinInlineOp} does.
 *
 * `expectedRetired` states the own fields of the paragraph the join retires,
 * and `expectedSurvivor` the review fields of the one it keeps, which the
 * join replaces; either refuses the join as stale when it differs (fields a
 * relayout recomputes are not compared).
 *
 * With `revision`, the join is tracked and moves nothing: the first
 * paragraph's mark becomes a tracked deletion, refused when the mark already
 * carries a tracked change, and the second takes the paragraph properties
 * the direct join would give it as a tracked property change. Accepting removes the mark, which
 * leaves the second paragraph with the first's content before its own: what
 * the direct join leaves. A tracked join always leaves the second, so
 * `survivor` must then be absent or `second`, and `depth` absent or zero: the
 * paragraph mark cannot record an inline merge depth.
 */
export type JoinBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.JOIN_BLOCKS;
  story: OpStory;
  blockId: string;
  nextBlockId: string;
  depth?: number;
  survivor?: SplitHalf;
  expectedRetired?: SplitParagraphFields;
  expectedSurvivor?: ParagraphReviewFields;
  /** Explicitly remove the first paragraph's section boundary with its mark. */
  sectionBoundary?: typeof SECTION_BOUNDARY_POLICIES.REMOVE;
  sectionView?: SectionViewChange;
  newIds?: NewIds;
  revision?: RevisionStamp;
};

/** A paragraph-relative insertion point in one block list. */
export type BlockInsertionPoint =
  | { type: "before"; blockId: string }
  | { type: "after"; blockId: string };

/**
 * Insert paragraphs beside an existing paragraph, with pre-minted ids.
 * Section-bearing paragraphs are refused. With `revision`, each inserted
 * paragraph receives an inserted mark and its content an insertion wrapper.
 * Tracked insertion requires a following paragraph in the same block list;
 * terminal insertion and insertion immediately before non-paragraph blocks
 * are refused as `untrackable`.
 * Incoming revision and content-control identities are refused in tracked
 * mode rather than nesting reviews or rewriting supplied identities.
 */
export type InsertBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_BLOCKS;
  story: OpStory;
  at: BlockInsertionPoint;
  blocks: readonly Paragraph[];
  revision?: RevisionStamp;
  newIds?: NewIds;
};

/**
 * Delete adjacent paragraphs in one block list. A following paragraph keeps
 * its identity and fields. At the end of a block list, the last selected
 * paragraph survives: it takes the preceding paragraph's content and
 * paragraph properties when present, keeping its own mark properties and id;
 * without a preceding paragraph it becomes empty.
 *
 * With `revision`, selected content is wrapped in deletions and the removed
 * breaks are marked. The container-final paragraph never receives a mark.
 * Section boundaries, non-paragraph following blocks and conflicting
 * paragraph reviews are refused rather than partially removed. Tracked deletion
 * of content carrying revision or content-control ids is also unsupported.
 */
export type DeleteBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_BLOCKS;
  story: OpStory;
  blockIds: readonly string[];
  revision?: RevisionStamp;
  newIds?: NewIds;
};

/**
 * Replace a run of adjacent paragraphs with other paragraphs.
 *
 * `expected` is the paragraphs as they stand: they are found by `paraId` and
 * compared structurally (without the fields a relayout recomputes), so a
 * replacement authored against another state is refused rather than applied
 * over it. Paragraph insertion uses an anchored replacement as its inverse.
 */
export type ReplaceBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.REPLACE_BLOCKS;
  story: OpStory;
  expected: readonly Paragraph[];
  blocks: readonly Paragraph[];
  /** Exact structural replacement may intentionally change section boundaries. */
  sectionBoundaries?: typeof SECTION_BOUNDARY_POLICIES.REPLACE;
  sectionView?: SectionViewChange;
};

/**
 * A paragraph's review fields: its paragraph properties, their tracked change
 * and its mark's tracked change. An absent field is absent on the paragraph.
 */
export type ParagraphReviewFields = {
  formatting?: ParagraphFormatting;
  propertyChanges?: ParagraphPropertyChange[];
  pPrMark?: ParagraphMarkChange;
};

/**
 * Set a paragraph's review fields exactly. `expected` is what they are now,
 * and `review` what they become; the inverse swaps the two. A new mark on a
 * paragraph that ends its container (the story body or a table cell) is
 * refused: there is no next paragraph to join it with.
 */
export type SetParagraphReviewOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW;
  story: OpStory;
  blockId: string;
  expected: ParagraphReviewFields;
  review: ParagraphReviewFields;
};

/**
 * Replace a paragraph's whole content. `expected` is the content as it
 * stands, compared structurally; the inverse swaps it with `content`. Ids the
 * new content carries must be unused elsewhere in the package.
 */
export type ReplaceInlineOp = {
  type: typeof DOCUMENT_OP_TYPES.REPLACE_INLINE;
  story: OpStory;
  blockId: string;
  expected: readonly ParagraphContent[];
  content: readonly ParagraphContent[];
};

/** What resolving a tracked change does with it. */
export const REVISION_DECISIONS = Object.freeze({ ACCEPT: "accept", REJECT: "reject" } as const);

/** One of {@link REVISION_DECISIONS}. */
export type RevisionDecision = (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS];

/**
 * Accept or reject tracked changes by revision id: insertions, deletions and
 * moves, run and paragraph property changes, paragraph marks, and row insertions
 * and deletions.
 *
 * - Accepting an insertion or rejecting a deletion keeps the content and
 *   drops the wrapper; the other two remove the content, changes nested in it
 *   included. A tracked change left with nothing in it goes.
 * - Accepting a property change drops the record; rejecting one restores the
 *   formatting it recorded.
 * - Accepting an inserted mark or rejecting a deleted one keeps the break and
 *   drops the record. The other two remove the mark, and the paragraph's
 *   properties go with it: the next paragraph is left, with its id,
 *   properties, mark and pending property changes, and the first's content
 *   before its own. The result does not depend on the order marks are
 *   resolved in. A mark with no next paragraph to join (a
 *   table follows, or the paragraph ends its container) loses its record,
 *   unless the paragraph holds no content and goes with it: removing a
 *   paragraph is a block operation, so that is refused (`untrackable`), as is
 *   a join at a section break.
 * - Row insertions and deletions are resolved after inline changes and before
 *   paragraph marks. Keeping a row clears its structural revision; removing
 *   it also removes its nested revisions. Removing every row removes the
 *   table, preserving the surrounding paragraphs.
 * - Records the resolution leaves meeting are merged as far as they are
 *   alike: the pieces a change cut apart are one record again.
 *
 * It is applied as the primitive operations it expands to, atomically, and
 * its inverse is theirs. Ids no record carries are skipped, so resolving the
 * same ids again changes nothing.
 */
export type ResolveRevisionOp = {
  type: typeof DOCUMENT_OP_TYPES.RESOLVE_REVISION;
  story: OpStory;
  revisionIds: readonly number[];
  decision: RevisionDecision;
};

/**
 * Insert a row at a zero-based index in the table holding `blockId`.
 * The paragraph reference selects the innermost table containing it.
 * With `revision`, the row, its cell content and
 * cell-ending paragraph marks record insertion. `newIds` supplies the
 * additional physical revision ids.
 */
export type InsertRowOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_ROW;
  story: OpStory;
  blockId: string;
  at: number;
  row: TableRow;
  revision?: RevisionStamp;
  newIds?: NewIds;
};

/**
 * Remove the row holding `blockId`, or record a row and cell-content deletion
 * with `revision`. The innermost containing row is selected. `expected`, when
 * present, refuses a stale row. Direct removal of the last row removes its
 * table; tracked deletion keeps a row not pending deletion.
 */
export type DeleteRowOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_ROW;
  story: OpStory;
  blockId: string;
  expected?: TableRow;
  revision?: RevisionStamp;
  newIds?: NewIds;
};

/**
 * Replace a table's rows exactly, with a structural staleness precondition.
 * This is the inverse and resolution primitive for row operations. An empty
 * row list removes the table, with an exact structural inverse.
 */
export type SetTableRowsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_TABLE_ROWS;
  story: OpStory;
  blockId: string;
  expected: readonly TableRow[];
  rows: readonly TableRow[];
};

/**
 * Insert a table before or after a paragraph. Tracking records each row,
 * cell content and cell-ending paragraph mark. At a container's end,
 * `terminal` moves the source content into a new preceding paragraph while
 * retaining the source id and fields on the untracked final paragraph.
 */
export type InsertTableOp = {
  type: typeof DOCUMENT_OP_TYPES.INSERT_TABLE;
  story: OpStory;
  at: BlockInsertionPoint;
  table: Table;
  revision?: RevisionStamp;
  newIds?: NewIds;
  terminal?: { beforeBlockId: string };
};

/**
 * Remove the innermost table holding the addressed paragraph. Its block list
 * must retain a paragraph for the inverse. Tracking records row, cell-content
 * and cell-ending paragraph-mark deletions without removing content.
 */
export type DeleteTableOp = {
  type: typeof DOCUMENT_OP_TYPES.DELETE_TABLE;
  story: OpStory;
  blockId: string;
  expected?: Table;
  revision?: RevisionStamp;
  newIds?: NewIds;
};

/**
 * Exact structural inverse for table edits. The addressed paragraph must
 * remain in the same block list; unrelated edits in that list make it stale.
 * Section boundaries and story-final marks are preserved; cell-ending marks
 * are exact model data, including after independent row resolution.
 */
export type SetContainerBlocksOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS;
  story: OpStory;
  blockId: string;
  expected: readonly BlockContent[];
  blocks: readonly BlockContent[];
};

/** Create a deterministically identified part and bind it to one section variant. */
export type CreateHeaderFooterOp = {
  type: typeof DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER;
  sectionIndex: number;
  story: HeaderFooterStory;
  referenceType: HeaderFooterType;
  content: BlockContent[];
};
export type RemoveHeaderFooterOp = {
  type: typeof DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER;
  sectionIndex: number;
  story: HeaderFooterStory;
  referenceType: HeaderFooterType;
};
export type AddNoteOp = {
  type: typeof DOCUMENT_OP_TYPES.ADD_NOTE;
  at: TextPosition;
  note: Footnote | Endnote;
};
export type RemoveNoteOp = {
  type: typeof DOCUMENT_OP_TYPES.REMOVE_NOTE;
  at: TextPosition;
  story: NoteStory;
};
export type SetSectionPropsOp = {
  type: typeof DOCUMENT_OP_TYPES.SET_SECTION_PROPS;
  sectionIndex: number;
  patch: FormattingPatch<SectionProperties>;
};
/** JSON-safe lifecycle deltas: omitted fields are unowned, null restores absence. */
export type StoryParts = {
  body?: {
    [Key in keyof Omit<DocumentBody, "sections">]?: Key extends "content"
      ? DocumentBody[Key]
      : Exclude<DocumentBody[Key], undefined> | null;
  };
  /** Section contents remain derived from the body's content and are never duplicated. */
  sections?: readonly {
    index: number;
    properties?: Section["properties"];
    headers?: readonly (readonly [HeaderFooterType, HeaderFooter])[] | null;
    footers?: readonly (readonly [HeaderFooterType, HeaderFooter])[] | null;
  }[];
  headers?: readonly (readonly [string, HeaderFooter])[] | null;
  footers?: readonly (readonly [string, HeaderFooter])[] | null;
  footnotes?: readonly Footnote[] | null;
  endnotes?: readonly Endnote[] | null;
  settings?: DocumentSettings | null;
};
/** Exact stale-checked lifecycle inverse; unrelated package parts stay untouched. */
export type RestoreStoryPartsOp = {
  type: typeof DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS;
  expected: StoryParts;
  parts: StoryParts;
};

/** A schema-version-5 document operation. */
export type DocumentOp =
  | CreateHeaderFooterOp
  | RemoveHeaderFooterOp
  | AddNoteOp
  | RemoveNoteOp
  | SetSectionPropsOp
  | RestoreStoryPartsOp
  | DeleteBlocksOp
  | InsertTableOp
  | DeleteTableOp
  | SetContainerBlocksOp
  | InsertBlocksOp
  | InsertTextOp
  | InsertContentOp
  | DeleteRangeOp
  | SplitInlineOp
  | JoinInlineOp
  | SetRunPropsOp
  | SetParagraphPropsOp
  | SplitBlockOp
  | JoinBlocksOp
  | ReplaceBlocksOp
  | SetParagraphReviewOp
  | ReplaceInlineOp
  | ResolveRevisionOp
  | InsertRowOp
  | DeleteRowOp
  | SetTableRowsOp;

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
