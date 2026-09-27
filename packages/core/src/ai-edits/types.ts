import type {
  FolioContentBlock,
  FolioContentInlineBooleanProperty,
  FolioContentInlineFormatting,
  FolioContentInlineFormattingPatch,
  FolioContentListReference,
  FolioContentParagraphIndentation,
  FolioContentParagraphSpacing,
  FolioContentParagraphKind,
  FolioContentRun,
  FolioContentSnapshot,
  FolioContentTableLocation,
} from "../compare/content-types";
import type { ListKind } from "../docx/listNumberingInstances";
import type { BreakContent, OutlineLevel, ParagraphAlignment } from "../types/document";

export type FolioAIBlockKind = FolioContentParagraphKind;

/** Boolean run properties supported by snapshots and range-formatting mutations. */
export type FolioAIInlineBooleanProperty = FolioContentInlineBooleanProperty;

export type FolioAIInlineFormatting = FolioContentInlineFormatting;

/**
 * A run-formatting mutation: `false` authors an explicit off value, while
 * `null` removes the direct property so its inherited value becomes effective.
 */
export type FolioAIInlineFormattingPatch = FolioContentInlineFormattingPatch;

export type FolioAIBlockPreviewRun = FolioContentRun;

/**
 * Inline structure inside a block's clean text.
 *
 * `pageBreak` is zero-width at one clean-text boundary and always represents
 * an authored `<w:br w:type="page"/>`; an omitted `clear` remains distinct from
 * any explicit value even though it does not affect layout for this break type.
 *
 * `noteReference` is a footnote or endnote reference, which the text shows as
 * its marker (`[^1]` for the first footnote, `[^e1]` for the first endnote —
 * the numbering the Markdown export uses) over `length` characters from
 * `offset`. The marker is structure, not text: a range may not cut into it,
 * and a text edit that would rewrite or remove it is refused
 * (`protectedReference`).
 */
export type FolioAIBlockStructuralBoundary =
  | {
      type: "pageBreak";
      offset: number;
      clear?: BreakContent["clear"];
    }
  | {
      type: "noteReference";
      noteType: "footnote" | "endnote";
      offset: number;
      length: number;
    };

/**
 * Where a block sits inside its innermost enclosing table. Every index is
 * zero-based: `tableIndex` counts tables in document order across the story,
 * `rowIndex` is the row's index in that table, `cellIndex` is the cell's
 * PHYSICAL index within the row (a merged cell occupies one slot, so this is
 * not a grid column), and `paragraphIndex` orders the block among the cell's
 * own blocks. Absent on a block that is not inside a table.
 */
export type FolioAIBlockTableLocation = FolioContentTableLocation;

export type FolioAIBlock = FolioContentBlock<FolioAIBlockKind> & {
  structuralBoundaries?: readonly FolioAIBlockStructuralBoundary[];
};

/**
 * The complete modeled attribute set of one direct `w:pPr/w:spacing` child.
 * Optional fields preserve the distinction between an absent attribute and
 * an explicit zero or false value.
 */
export type FolioAIParagraphSpacing = FolioContentParagraphSpacing;

/** The complete modeled attribute set of one direct `w:pPr/w:ind` child. */
export type FolioAIParagraphIndentation = FolioContentParagraphIndentation;
export type FolioAIListReference = FolioContentListReference;

/** The kinds of list an operation can start. */
export type FolioAIListKind = ListKind;

/**
 * Start a new list: a numbering instance of `kind`, defined for the
 * operation, with its paragraphs at `level` (zero-based, default 0). Every
 * paragraph one operation numbers this way joins the same new list; separate
 * operations start separate lists. A package without a numbering part gets
 * one.
 */
export type FolioAINewListReference = {
  start: "new";
  kind: FolioAIListKind;
  level?: number;
};

/** An existing numbering instance and level, or a new list. */
export type FolioAIListNumbering = FolioAIListReference | FolioAINewListReference;

/** Which paragraphs split from one insertion's `text` receive its paragraph formatting. */
export type FolioAIInsertFormattingScope = "firstParagraph" | "allParagraphs";

/**
 * The paragraph properties an operation may set. A subset of `w:pPrChange`'s
 * scope: properties a comparison can see in a block projection and an agent
 * has a reason to change.
 */
export type FolioAIBlockParagraphProperties = {
  /** `w:pStyle`. `null` clears the style back to the default. */
  styleId?: string | null;
  /** Direct `w:outlineLvl`. `null` restores style inheritance. */
  outlineLevel?: OutlineLevel | null;
  /**
   * `w:numPr/w:ilvl`, zero-based. `null` removes numbering unless `numbering`
   * supplies a concrete instance; together they retain that instance without
   * an authored level.
   */
  listLevel?: number | null;
  numbering?: FolioAIListNumbering | null;
  /** Direct `w:jc`. `null` clears the override and restores style inheritance. */
  alignment?: ParagraphAlignment | null;
  /** Direct `w:spacing` attributes. `null` removes the whole direct child. */
  spacing?: FolioAIParagraphSpacing | null;
  /** Direct `w:ind` attributes. `null` removes the whole direct child. */
  indentation?: FolioAIParagraphIndentation | null;
};

/**
 * Every paragraph of one story, blank ones included.
 *
 * A blank paragraph is part of a document's shape: an empty cell is a column,
 * an empty row is a row, and an operation or a comparison that cannot address
 * them cannot describe what changed around them. A surface that reads the
 * document for a person or a model wants only the paragraphs that carry text,
 * and says so with `isFolioAIContentBlock`.
 */
export type FolioAIEditSnapshot = FolioContentSnapshot<FolioAIBlock> & {
  anchors: Record<string, FolioAIBlockAnchor>;
};

export type FolioAIBlockAnchor = {
  id: string;
  from: number;
  to: number;
  text: string;
  normalizedText: string;
  textHash: string;
  /** Hash of zero-width structural boundaries in the clean block view. */
  structuralBoundaryHash: string;
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
        formatting: FolioAIInlineFormattingPatch;
      }
    | {
        id: string;
        type: "insertAfterBlock" | "insertBeforeBlock";
        blockId: string;
        /**
         * The paragraph text to insert. With the default `lineBreakMode:
         * "paragraph"`, a line break splits `text` into consecutive
         * paragraphs at the same anchor. `formattingScope` decides which of
         * them receive this operation's paragraph formatting: by default only
         * the first, later ones use body formatting. Blank lines are dropped
         * and reported as a `splitMultilineText` normalization. `"inline"`
         * retains the control inside one paragraph.
         *
         * `""` inserts a BLANK paragraph, and is a real edit: adding an empty
         * line is a change a reader sees, and a document that has one where
         * another does not differs. It is applied like any other insertion,
         * with a tracked paragraph mark, so rejecting closes it away.
         */
        text: string;
        /**
         * `"paragraph"` (the default) splits newlines into consecutive blocks.
         * `"inline"` retains tabs and hard breaks inside this inserted block.
         */
        lineBreakMode?: "paragraph" | "inline";
        /**
         * Which paragraphs split from `text` receive the operation's
         * paragraph formatting: the inherited anchor formatting plus
         * `styleId`, `listLevel`, `numbering`, `alignment`, `spacing` and
         * `indentation`. `"firstParagraph"` (the default) formats the first
         * and leaves the rest as body paragraphs, for a heading followed by
         * its body. `"allParagraphs"` formats every paragraph alike, for
         * several list items in one operation. `pageBreakBefore` and
         * `hardPageBreak` apply to the first paragraph in either scope.
         */
        formattingScope?: FolioAIInsertFormattingScope;
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
         * Insert one standalone authored `<w:br w:type="page"/>` carrier.
         * This differs from `pageBreakBefore`, which is a paragraph property.
         */
        hardPageBreak?: {
          clear?: BreakContent["clear"];
        };
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
         * beside a list item at a different depth. `null` gives it no
         * numbering unless `numbering` supplies a concrete instance; together
         * they retain that instance without an authored level.
         */
        listLevel?: number | null;
        numbering?: FolioAIListNumbering | null;
        /**
         * Direct `w:jc` for the inserted block. `null` clears alignment copied
         * from the anchor and lets the inserted paragraph's style decide.
         */
        alignment?: ParagraphAlignment | null;
        /**
         * Direct `w:spacing` for the inserted block. `null` clears spacing
         * copied from the anchor and lets the inserted paragraph's style decide.
         */
        spacing?: FolioAIParagraphSpacing | null;
        /** Direct `w:ind` for the inserted block. `null` clears copied indentation. */
        indentation?: FolioAIParagraphIndentation | null;
        comment?: FolioAIComment;
      }
    | {
        id: string;
        type: "replaceBlock";
        blockId: string;
        text: string;
        preserveFormatting?: boolean;
        /** Paragraph style to apply; `null` clears the direct style. */
        styleId?: string | null;
        comment?: FolioAIComment;
      }
    /**
     * Delete the whole block. A block with words loses them and its paragraph
     * mark; a BLANK block has only a paragraph mark to lose, and loses it, so
     * the empty line goes away rather than the operation doing nothing.
     *
     * The block that ENDS its container is the exception, in both modes: a
     * body, a cell, a header or footer, a note and a text box each end with a
     * paragraph, and a deleted mark there would say "join with the paragraph
     * after this one" where there is none. It loses its words and keeps its
     * place. To remove it, delete the mark of the block BEFORE it, which
     * merges forward into it.
     */
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
        /**
         * Paragraph properties for the first result. Omitted properties keep
         * the source paragraph's value.
         */
        firstParagraphProperties?: FolioAIBlockParagraphProperties;
        /**
         * Paragraph properties for the second result. Omitted properties keep
         * the source paragraph's value.
         */
        secondParagraphProperties?: FolioAIBlockParagraphProperties;
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
        /**
         * Cell texts row by row. Every row must hold the same number of cells.
         *
         * A cell holds paragraphs, not lines: a line break in a cell's text
         * starts a new paragraph in that cell, and a blank line is a blank
         * paragraph. That is not the rule `insertAfterBlock` follows for
         * prose, where a blank line between two model-written clauses is
         * formatting noise and is dropped — a cell's text describes paragraphs
         * that exist, so dropping one would lose a block.
         */
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
        /**
         * Paragraph properties for the joined result. Omitted properties keep
         * the first paragraph's value.
         */
        mergedParagraphProperties?: FolioAIBlockParagraphProperties;
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
        /**
         * Initial text for each cell of the new row, in order — the same order
         * a block's `table.cellIndex` counts. Cells left unnamed stay empty.
         *
         * Sized against the cells the new row really has, which is fewer than
         * the table's columns where a vertical merge from the row above
         * reaches through the insertion point (the merge grows through the
         * new row and keeps its column). More texts than cells is refused as
         * `payloadDoesNotFit`, naming the texts with no cell, before anything
         * is written; fewer leaves the remaining cells empty.
         *
         * A line break in a cell's text starts a new paragraph in that cell,
         * as in `insertTable`.
         */
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
        /**
         * Initial text for newly created physical cells in row order. A
         * horizontal merge the new column crosses widens instead of taking a
         * cell, so more texts than new cells is refused as `payloadDoesNotFit`.
         */
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
 *   formatRange), paragraph properties, plus the block and table row/column structural
 *   operations (insertAfterBlock, insertBeforeBlock, replaceBlock, deleteBlock,
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
  /** The paragraph already owns the one `w:pPrChange` OOXML permits. */
  | "pendingParagraphPropertyChange"
  /** The affected run already owns the one `w:rPrChange` OOXML permits. */
  | "pendingRunPropertyChange"
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
  | "documentNotEditable"
  /**
   * The operation supplies more values than its target can hold, so applying
   * it would drop some of them: an `insertTableRow` whose `cellTexts` outnumber
   * the cells the new row has (a vertical merge crossing the insertion point
   * extends through the new row and takes a column away from it), or an
   * `insertTableColumn` whose `cellTexts` outnumber the cells the new column
   * has. Nothing was applied; `message` names the values that do not fit.
   * Supply fewer values, or anchor the operation where the table has room.
   */
  | "payloadDoesNotFit"
  /**
   * An earlier operation of the same batch already claims this operation's
   * target: it deletes, rewrites, splits or merges the block this one edits,
   * or edits an overlapping stretch of its text. Operations of one batch all
   * address the document as it was read, so this one would have landed on
   * positions the other had moved. Nothing of it was applied; re-read the
   * document after the batch and send it again, on its own.
   */
  | "overlappingOperation"
  /**
   * An offset of the operation falls inside one character: between the two
   * UTF-16 halves of a surrogate pair (an emoji, a character outside the
   * Basic Multilingual Plane), or — for an operation that changes text or
   * breaks a paragraph — inside a grapheme cluster (a letter and its combining
   * marks, an emoji sequence joined with zero-width joiners, a flag). Text cut
   * there cannot be written back whole. Nothing was applied; move the offset
   * to the boundary before or after the character.
   */
  | "splitsCharacter"
  /**
   * The text change would rewrite or remove a footnote or endnote reference
   * (the `[^1]` / `[^e1]` markers the text shows), or write a marker-shaped
   * string of its own. A reference is structure, not text: keep each marker
   * the match covers in the replacement, in order, or match only the prose
   * beside it.
   */
  | "protectedReference"
  /**
   * The operation names a paragraph style (`styleId`) the document does not
   * define, or one defined as a table, character or numbering style. Nothing
   * was applied: such a `w:pStyle` confers no formatting, so the paragraph
   * would keep its body look. Use a paragraph style the document defines, or
   * `null` to clear the style.
   */
  | "missingStyle"
  /**
   * The block is pending deletion: a tracked change deletes its text and its
   * paragraph mark, or the table row or cell holding it, so a reader lists it
   * as a blank block and accepting removes it. Nothing was applied: text
   * written there would join the next paragraph, or go with the row, once the
   * deletion is accepted. Reject that deletion first, or insert a new block
   * next to it.
   */
  | "pendingDeletion";

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
   * The list also includes a paragraph-property revision synthesized
   * when an inserted final paragraph mark rotates to its carrier.
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
  /**
   * What exactly was wrong when the reason alone does not say: the values a
   * `payloadDoesNotFit` skip could not place, the earlier operation an
   * `overlappingOperation` skip conflicts with, the offset and character of a
   * `splitsCharacter` skip, or the undefined style of a `missingStyle` skip.
   */
  message?: string;
};

/**
 * One automatic adjustment `apply.ts` made to an operation's input to keep the
 * applied result well-formed. Reported rather than applied silently: the
 * caller asked for something the document could not hold, and gets told what
 * it got instead.
 */
export type FolioAIEditNormalization =
  /**
   * A line-break in paragraph-mode `insertAfterBlock` /
   * `insertBeforeBlock` text was split into one paragraph per non-blank line.
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
  | { id: string; code: "unpairedMove"; moveId: string }
  /**
   * A text replacement rewrote a stretch whose characters did not all carry
   * the same formatting, link or comment. The new text takes the formatting
   * of the first character it replaces, and every link and comment over the
   * stretch; text the replacement keeps keeps its own. A change spanning
   * several words keeps only whole words, so no word is left partly in the
   * old formatting. Direct and tracked modes allocate alike: accepting the
   * tracked replacement leaves what the direct one writes.
   */
  | { id: string; code: "uniformReplacementFormatting" };

export type FolioAIEditNormalizationCode = FolioAIEditNormalization["code"];

export type FolioAIEditApplyResult = {
  applied: FolioAIEditAppliedOperation[];
  skipped: FolioAIEditSkippedOperation[];
  /** Present only when at least one operation triggered a normalization. */
  normalizations?: FolioAIEditNormalization[];
};
