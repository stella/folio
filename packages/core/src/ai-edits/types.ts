export type FolioAIBlockKind = "heading" | "listItem" | "paragraph";

export type FolioAIBlockPreviewRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
};

/**
 * Where a block sits inside its innermost enclosing table. Every index is
 * zero-based: `tableIndex` counts tables in document order across the story,
 * `rowIndex` is the row's index in that table, `cellIndex` is the cell's
 * PHYSICAL index within the row (a merged cell occupies one slot, so this is
 * not a grid column), and `paragraphIndex` orders the block among the cell's
 * own blocks. Absent on a block that is not inside a table.
 */
export type FolioAIBlockTableLocation = {
  /**
   * Document-order index of the OUTERMOST table the block sits in — the same
   * as `tableIndex` unless tables nest. A comparison aligns on this: a table
   * inside a cell is part of its parent, not a structure of its own that can
   * be paired against one somewhere else.
   */
  outerTableIndex: number;
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
  paragraphIndex: number;
};

export type FolioAIBlock = {
  id: string;
  kind: FolioAIBlockKind;
  text: string;
  /** One-based heading depth when the block has outline semantics. */
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
  /**
   * `w:numPr/w:ilvl`: the block's list indent level. Present only on a block
   * that carries numbering, and the only pPr property a redline can move
   * without touching a word — a demoted list item reads as unchanged text and
   * is not.
   */
  listLevel?: number;
  previewRuns?: FolioAIBlockPreviewRun[];
  table?: FolioAIBlockTableLocation;
};

/**
 * The paragraph properties an operation may set. A subset of `w:pPrChange`'s
 * scope: the two a comparison can see in a block projection, and the two an
 * agent has a reason to change.
 */
export type FolioAIBlockParagraphProperties = {
  /** `w:pStyle`. `null` clears the style back to the default. */
  styleId?: string | null;
  /** `w:numPr/w:ilvl`, zero-based. */
  listLevel?: number;
};

export type FolioAIEditSnapshot = {
  blocks: FolioAIBlock[];
  anchors: Record<string, FolioAIBlockAnchor>;
  /** Hidden empty paragraph used to anchor insertions when `blocks` is empty. */
  emptyDocumentAnchorId?: string;
};

export type FolioAIBlockAnchor = {
  id: string;
  from: number;
  to: number;
  text: string;
  normalizedText: string;
  textHash: string;
  hashOccurrenceCount: number;
};

export type FolioAIComment = {
  text: string;
};

export type FolioAIEditSeverity = "low" | "medium" | "high";

/**
 * Optional review metadata attached to an AI-authored operation.
 * Set by the model when performing a structured review (e.g.
 * `severity: "high"`, `area: "Penalty"`); absent for direct edits.
 * Both fields are independent — either or both may be set.
 */
export type FolioAIEditReviewMeta = {
  severity?: FolioAIEditSeverity;
  area?: string;
};

export type FolioAIEditPrecondition = {
  blockTextHash: string;
};

/**
 * A serializable range over the visible, post-tracked-changes text of one
 * block. Offsets are zero-based UTF-16 boundaries, matching JavaScript string
 * slicing; `selectedTextHash` makes a shifted or changed selection fail stale.
 */
export type FolioAITextRangeHandle = {
  type: "textRange";
  story: "main";
  blockId: string;
  startOffset: number;
  endOffset: number;
  selectedTextHash: string;
};

/**
 * Stable handle for the logical document section introduced by one heading.
 * The section runs until the next heading at the same or a higher level.
 * `headingTextHash` makes a renamed heading fail stale instead of silently
 * resolving to content whose meaning may have changed.
 */
export type FolioDocumentSectionHandle = {
  type: "headingSection";
  story: "main";
  headingBlockId: string;
  headingTextHash: string;
  /** One-based depth used to detect structural section-boundary changes. */
  headingLevel: number;
};

export type FolioDocumentOutlineEntry = {
  handle: FolioDocumentSectionHandle;
  headingBlockId: string;
  text: string;
  /** One-based heading depth. */
  level: number;
  parentHandle?: FolioDocumentSectionHandle;
};

export type FolioDocumentOutline = {
  sections: FolioDocumentOutlineEntry[];
};

export type FolioDocumentSection = {
  handle: FolioDocumentSectionHandle;
  heading: FolioDocumentOutlineEntry;
  /** Heading block followed by every block in its logical section. */
  blocks: FolioAIBlock[];
};

export type FolioDocumentSectionReadResult =
  | { status: "found"; section: FolioDocumentSection }
  | { status: "missing" }
  | { status: "stale" };

export type FolioDocumentNavigationTarget =
  | { type: "block"; story: "main"; blockId: string }
  | FolioAITextRangeHandle;

export type FolioAIInlineFormatting = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

/**
 * A party in an `insertSignatureTable` op. Mirrors the
 * `signatureTable` helper in `docx-core/legal-source/compile.ts`:
 * name is rendered bold; `signatory` and `title` are optional
 * lines under the signature rule (title in italics).
 */
export type FolioAISignatureParty = {
  name: string;
  signatory?: string;
  title?: string;
};

export type FolioAIEditOperation = FolioAIEditReviewMeta & {
  precondition?: FolioAIEditPrecondition;
  /**
   * Groups this operation's produced marks under one logical suggestion so the
   * host can accept/reject them together. Only consulted in `"suggested"` apply
   * mode; ignored otherwise. When omitted in suggested mode the applier falls
   * back to the operation `id`, giving per-operation grouping.
   */
  suggestionId?: string;
} & (
    | {
        id: string;
        type: "replaceInBlock";
        blockId: string;
        find: string;
        replace: string;
        comment?: FolioAIComment;
      }
    | {
        id: string;
        type: "replaceRange";
        range: FolioAITextRangeHandle;
        replace: string;
        comment?: FolioAIComment;
      }
    | {
        id: string;
        type: "commentOnRange";
        range: FolioAITextRangeHandle;
        comment: FolioAIComment;
      }
    | {
        id: string;
        type: "formatRange";
        range: FolioAITextRangeHandle;
        formatting: FolioAIInlineFormatting;
      }
    | {
        id: string;
        type: "insertAfterBlock" | "insertBeforeBlock";
        blockId: string;
        /**
         * The paragraph text to insert. A line break splits `text` into
         * consecutive paragraphs at the same anchor instead of becoming
         * literal newlines inside one paragraph: only the first paragraph
         * gets `styleId` / `inheritFormatting`, later ones use body
         * formatting. Blank lines are dropped. Reported as a
         * `splitMultilineText` normalization when it happens.
         */
        text: string;
        inheritFormatting?: boolean;
        /**
         * Links this insertion to the deletion that carries the same
         * `moveId`: together they are one relocation, written as `w:moveTo`
         * and `w:moveFrom`. See `deleteBlock`.
         */
        moveId?: string;
        /**
         * When true, mark the inserted paragraph with
         * `pageBreakBefore` so the layout engine starts it on a
         * new page. Use for explicit page-break inserts.
         */
        pageBreakBefore?: boolean;
        /**
         * Override the paragraph `styleId` attr of the inserted
         * block (e.g. `ClauseHeading1`). When omitted the inserted
         * block inherits the source block's styleId via
         * `inheritFormatting`; `null` gives it no style at all, which
         * inheritance alone cannot say.
         */
        styleId?: string | null;
        /**
         * Override `w:numPr/w:ilvl` on the inserted block, keeping the
         * anchor's `w:numId`. Without it the inserted paragraph takes the
         * anchor's level, which is the wrong one whenever the new item sits
         * beside a list item at a different depth.
         */
        listLevel?: number;
        comment?: FolioAIComment;
      }
    | {
        id: string;
        type: "replaceBlock";
        blockId: string;
        text: string;
        preserveFormatting?: boolean;
        styleId?: string;
        comment?: FolioAIComment;
      }
    | {
        id: string;
        type: "deleteBlock";
        blockId: string;
        /**
         * Links this deletion to the insertion that carries the same
         * `moveId`: together they are one relocation, written as `w:moveFrom`
         * and `w:moveTo` rather than as an unrelated deletion and insertion.
         * A `moveId` that does not name exactly one of each is reported as an
         * `unpairedMove` normalization and both halves apply plainly.
         */
        moveId?: string;
        comment?: FolioAIComment;
      }
    /**
     * Break the block in two at `offset`, moving a paragraph mark and no
     * words. In tracked-changes mode the first half carries an INSERTED
     * paragraph mark, so accepting keeps the break and rejecting closes it;
     * the alternative — rewriting the first half and inserting the second —
     * claims the tail was newly written when nobody touched it.
     */
    | {
        id: string;
        type: "splitBlock";
        /** Offset in the block's text. Must fall strictly inside it. */
        offset: number;
        /**
         * Text at `offset` the break replaces — the space between the two
         * halves, when the split consumed one. Deleted in `"direct"` mode and
         * deletion-marked in tracked mode, so rejecting restores it.
         */
        separator?: string;
        blockId: string;
      }
    /**
     * Add a whole table next to the anchor block, its rows marked inserted in
     * tracked mode. `insertTableRow` can only grow a table that already
     * exists; a comparison whose target gained one needs to say so.
     */
    | {
        id: string;
        type: "insertTable";
        blockId: string;
        /** Place the table after the anchor (default) or before it. */
        position?: "after" | "before";
        /** Cell texts row by row. Every row must hold the same number of cells. */
        rows: readonly (readonly string[])[];
      }
    /**
     * Remove the whole table the block sits in, its rows marked deleted in
     * tracked mode. The mirror of `insertTable`.
     */
    | {
        id: string;
        type: "deleteTable";
        blockId: string;
      }
    /**
     * Replace the block's paragraph properties, recorded as a `w:pPrChange`
     * in tracked mode so the previous set is restored on reject. The edit
     * that moves no words: a list item demoted a level, a paragraph restyled
     * as a heading.
     */
    | {
        id: string;
        type: "setBlockParagraphProperties";
        blockId: string;
        properties: FolioAIBlockParagraphProperties;
      }
    /**
     * Join the block with the one after it, the mirror of `splitBlock`: in
     * tracked-changes mode the block carries a DELETED paragraph mark, so
     * accepting closes the break and rejecting keeps it.
     *
     * Refused when the block has no joinable sibling — the last paragraph of
     * a table cell, or of the story — because a deleted mark there would
     * accept into a join that cannot happen and leave a revision no reader
     * can resolve.
     */
    | {
        id: string;
        type: "mergeBlockWithNext";
        /**
         * Text the join inserts between the two halves — the space the
         * paragraph break used to stand in for. Insertion-marked in tracked
         * mode, so rejecting removes it along with the join.
         */
        separator?: string;
        blockId: string;
      }
    | {
        id: string;
        type: "commentOnBlock";
        blockId: string;
        quote?: string;
        comment: FolioAIComment;
      }
    | {
        id: string;
        type: "insertSignatureTable";
        blockId: string;
        /**
         * Position the table after the anchor block (default) or
         * before it. Always inserts as a sibling at the document
         * level — no nested-table support.
         */
        position?: "after" | "before";
        parties: FolioAISignatureParty[];
        comment?: FolioAIComment;
      }
    | {
        id: string;
        type: "insertTableRow";
        /** Stable paragraph anchor inside the row that receives the new sibling. */
        blockId: string;
        position?: "after" | "before";
        /** Initial text for each physical cell in source order; omitted cells stay empty. */
        cellTexts?: string[];
      }
    | {
        id: string;
        type: "deleteTableRow";
        /** Stable paragraph anchor inside the row to delete. */
        blockId: string;
      }
    | {
        id: string;
        type: "insertTableColumn";
        /** Stable paragraph anchor inside the cell that receives the new sibling column. */
        blockId: string;
        position?: "after" | "before";
        /** Initial text for newly created physical cells in row order. */
        cellTexts?: string[];
      }
    | {
        id: string;
        type: "deleteTableColumn";
        /** Stable paragraph anchor inside the column to delete. */
        blockId: string;
      }
    | ({
        id: string;
        type: "mergeTableCells";
        /** Stable paragraph anchor inside the first cell. */
        blockId: string;
      } & (
        | {
            /** Stable paragraph anchor inside the opposite corner cell. */
            endBlockId: string;
            rowCount?: never;
          }
        | {
            /** Number of grid rows to merge downward from the anchored cell. */
            rowCount: number;
            endBlockId?: never;
          }
      ))
    | {
        id: string;
        type: "splitTableCell";
        /** Stable paragraph anchor inside the cell to split. */
        blockId: string;
      }
  );

/**
 * How AI-authored operations are applied:
 * - `"direct"` — edits are written straight into the document.
 * - `"tracked-changes"` — edits land as normal tracked changes (`w:ins`/`w:del`).
 * - `"suggested"` — edits land as tracked changes carrying `"suggested"`
 *   provenance: rendered with the tracked-change grammar but stripped from
 *   serialized DOCX until a human accepts them. Behaves like `"tracked-changes"`
 *   and covers the inline text/format operations (replaceInBlock, replaceRange,
 *   formatRange) plus the block and table row/column structural operations
 *   (insertAfterBlock, insertBeforeBlock, replaceBlock, deleteBlock,
 *   insertSignatureTable, insertTableRow, deleteTableRow, insertTableColumn,
 *   deleteTableColumn). Only comment operations and cell merge/split report
 *   `unsupportedMode`.
 */
export type FolioAIEditApplyMode = "direct" | "tracked-changes" | "suggested";

export type FolioAIEditSkipReason =
  | "missingBlock"
  | "changedBlock"
  | "ambiguousFind"
  | "missingFind"
  | "unsupportedBlock"
  | "unsupportedMode"
  | "atomicBatchRejected"
  | "preconditionFailed"
  | "staleRange"
  | "emptyOperation"
  /**
   * The operation would not change the document — find equals
   * replace, or replaceBlock's `text` matches the live block.
   * Filtered out so the reviewer doesn't see "X → X" cards.
   */
  | "noopOperation"
  /**
   * The batch carried a `precondition.documentVersion` that no longer
   * matches the document the surface holds; every operation in the batch
   * is skipped. Re-read the document and regenerate the edits.
   */
  | "documentVersionMismatch"
  /**
   * The surface holds no editable document right now (no editor mounted,
   * no entity to attach suggestions to); the operation was neither applied
   * nor queued.
   */
  | "documentNotEditable";

export type FolioAIEditAppliedOperation = {
  id: string;
  commentId?: number;
  /**
   * Primary tracked-change revision id (only set when applied in
   * `tracked-changes` mode and the operation produced at least one
   * insertion/deletion mark). Stable identifier suitable for
   * scroll-to and visual reference.
   */
  revisionId?: number;
  /**
   * Every revision id this operation produced. A replace allocates
   * separate ids for the deletion and the insertion sides because
   * fromProseDoc serialises a single id carrying both as a Word
   * "moveTo/moveFrom" pair, not an ins/del — so the two sides must
   * be distinct ids in the doc but conceptually one operation here.
   * Use this list when you need to accept or reject every mark
   * belonging to this op.
   */
  revisionIds?: readonly number[];
  /**
   * The suggestion id stamped on every mark this operation produced (only set
   * in `"suggested"` apply mode). Pass it to `acceptSuggestion` /
   * `rejectSuggestion` / `scrollToSuggestion` to resolve the whole suggestion.
   */
  suggestionId?: string;
};

export type FolioAIEditSkippedOperation = {
  id: string;
  reason: FolioAIEditSkipReason;
};

/**
 * One automatic adjustment `apply.ts` made to an operation's input to keep the
 * applied result well-formed. Reported rather than applied silently: the
 * caller asked for something the document could not hold, and gets told what
 * it got instead.
 */
export type FolioAIEditNormalization =
  /**
   * A line-break in `insertAfterBlock` / `insertBeforeBlock`'s `text` cannot
   * become one paragraph with an embedded break (Word paragraphs are single
   * lines); the applier split it into one paragraph per non-blank line.
   */
  | {
      id: string;
      code: "splitMultilineText";
      /** Number of paragraphs the operation's `text` was split into. */
      paragraphCount: number;
    }
  /**
   * A `moveId` that did not name exactly one deletion and one insertion in
   * the batch. The operation still applies, as an ordinary insertion or
   * deletion: half a move pair is not a move, and `w:moveTo` without its
   * `w:moveFrom` is a relocation from nowhere.
   */
  | { id: string; code: "unpairedMove"; moveId: string };

export type FolioAIEditNormalizationCode = FolioAIEditNormalization["code"];

export type FolioAIEditApplyResult = {
  applied: FolioAIEditAppliedOperation[];
  skipped: FolioAIEditSkippedOperation[];
  /** Present only when at least one operation triggered a normalization. */
  normalizations?: FolioAIEditNormalization[];
};
