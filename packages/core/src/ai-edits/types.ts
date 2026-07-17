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

export type FolioAIBlock = {
  id: string;
  kind: FolioAIBlockKind;
  text: string;
  /** One-based heading depth when the block has outline semantics. */
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
  previewRuns?: FolioAIBlockPreviewRun[];
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
        text: string;
        inheritFormatting?: boolean;
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
         * `inheritFormatting`.
         */
        styleId?: string;
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
        comment?: FolioAIComment;
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

export type FolioAIEditApplyMode = "direct" | "tracked-changes";

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
  | "noopOperation";

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
};

export type FolioAIEditSkippedOperation = {
  id: string;
  reason: FolioAIEditSkipReason;
};

export type FolioAIEditApplyResult = {
  applied: FolioAIEditAppliedOperation[];
  skipped: FolioAIEditSkippedOperation[];
};
