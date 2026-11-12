/** Public batches compile against the read projection; only the canonical journal publishes. */
import {
  panic,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";
import { Mark, type Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import {
  CONTENT_CONTROL_LOCK,
  lockedControlMessage,
  paragraphNumberingReference,
  type Table,
} from "@stll/docx-core/model";
import {
  applyDocumentOps,
  type BlockInsertionPoint,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  documentStories,
  type EditorIntent,
  findStoryBody,
  freshCommentId,
  type IdLease,
  type NewListInstance,
  type OpStory,
  paragraphLength,
  type ParagraphPropsPatch,
  pendingSuggestionIdsIn,
  revisionSpan,
  type RunPropsPatch,
  storyParagraphs,
  type TextPosition,
} from "@stll/docx-core/ops";

import {
  buildTableNode,
  SUGGESTED_SUPPORTED_OPERATION_TYPES,
} from "../ai-edits/apply";
import {
  splitsGraphemeCluster,
  splitsSurrogatePair,
} from "../ai-edits/character-boundaries";
import {
  buildCleanBlockText,
  resolveCleanTextRange,
} from "../ai-edits/clean-text";
import { stripInlineEmphasisMarkers } from "../ai-edits/inline-emphasis";
import {
  createBatchListMinter,
  type MintedList,
} from "../ai-edits/newListNumbering";
import type { NoteReferenceLabels } from "../ai-edits/note-references";
import {
  isDeletedPPrMark,
  joinsIn,
  pendingJoinParagraphs,
} from "../ai-edits/pendingJoin";
import {
  collectNoteReferenceLabels,
  hashFolioAIBlockContent,
  hashFolioAIBlockText,
} from "../ai-edits/snapshot";
import { findOutermostTableBoundary } from "../ai-edits/table-targets";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditNormalization,
  FolioAIEditSkippedOperation,
  FolioAIEditSkipReason,
  FolioAIListKind,
  FolioAIListNumbering,
  FolioAIParagraphIndentation,
  FolioAIParagraphSpacing,
} from "../ai-edits/types";
import { wordDiffSessionFromOptions } from "../ai-edits/word-diff";
import type { FolioContentStatedNumbering } from "../compare/content-types";
import {
  type ApplyFolioDocumentOperationsOptions,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationStory,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
  getFolioDocumentOperationIssues,
  getFolioDocumentOperationReceiptsForStory,
  parseFolioDocumentOperationBatch,
  withDistinctSuggestionIds,
} from "../document-operations";
import { expectTableAttrs } from "../prosemirror/attrs";
import {
  standaloneTableRowsFromProseMirror,
  tableAttrsToFormatting,
} from "../prosemirror/conversion/fromProseDoc";
import { maxAnnotationIdInDoc } from "../prosemirror/plugins/revisionIds";
import {
  CANONICAL_GAP,
  type CanonicalGap,
} from "../types/canonicalCapabilities";
import type {
  Document,
  Paragraph,
  ParagraphFormatting,
  Run,
} from "../types/document";
import { readModel } from "../types/documentModel";
import {
  canonicalCommentBody,
  compileCanonicalComments,
} from "./canonicalComments";
import { lockViolationBetween } from "./canonicalLockGuard";
import {
  type CanonicalCommit,
  type CanonicalProjection,
  type CanonicalSession,
} from "./canonicalSession";

type CanonicalPublicOperationRefusal = NonNullable<
  FolioAIEditSkippedOperation["canonicalRefusal"]
>;

/** A new public kind cannot bypass a canonical compiler/refusal decision. */
export const CANONICAL_PUBLIC_OPERATION_DISPOSITIONS: {
  readonly replaceInBlock: "compile";
  readonly replaceRange: "compile";
  readonly replaceBlock: "compile";
  readonly splitBlock: "compile";
  readonly formatRange: "compile";
  readonly mergeBlockWithNext: "compile";
  readonly setBlockParagraphProperties: "compile";
  readonly insertAfterBlock: "compile";
  readonly insertBeforeBlock: "compile";
  readonly deleteBlock: "compile";
  readonly commentOnBlock: "compile";
  readonly commentOnRange: "compile";
  readonly insertTable: "compile";
  readonly insertSignatureTable: "publicOps.tableStructure";
  readonly deleteTable: "publicOps.tableStructure";
  readonly insertTableRow: "publicOps.tableStructure";
  readonly deleteTableRow: "publicOps.tableStructure";
  readonly insertTableColumn: "publicOps.tableStructure";
  readonly deleteTableColumn: "publicOps.tableStructure";
  readonly mergeTableCells: "publicOps.tableStructure";
  readonly splitTableCell: "publicOps.tableStructure";
} = {
  replaceInBlock: "compile",
  replaceRange: "compile",
  replaceBlock: "compile",
  splitBlock: "compile",
  formatRange: "compile",
  mergeBlockWithNext: "compile",
  setBlockParagraphProperties: "compile",
  insertAfterBlock: "compile",
  insertBeforeBlock: "compile",
  deleteBlock: "compile",
  commentOnBlock: "compile",
  commentOnRange: "compile",
  insertTable: "compile",
  insertSignatureTable: CANONICAL_GAP.publicTableStructure,
  deleteTable: CANONICAL_GAP.publicTableStructure,
  insertTableRow: CANONICAL_GAP.publicTableStructure,
  deleteTableRow: CANONICAL_GAP.publicTableStructure,
  insertTableColumn: CANONICAL_GAP.publicTableStructure,
  deleteTableColumn: CANONICAL_GAP.publicTableStructure,
  mergeTableCells: CANONICAL_GAP.publicTableStructure,
  splitTableCell: CANONICAL_GAP.publicTableStructure,
} as const satisfies Record<
  FolioDocumentOperation["type"],
  "compile" | CanonicalGap
>;

export type CanonicalPublicOperationOptions = Omit<
  ApplyFolioDocumentOperationsOptions,
  "view" | "createUndoHandle" | "createCommentId"
>;

type CommentOperation = {
  [
    Operation in FolioDocumentOperation as Operation["type"]
  ]: "comment" extends keyof Operation ? Operation : never;
}[FolioDocumentOperation["type"]];

type ResolvedOperation = {
  intents: readonly EditorIntent[];
  normalizations: readonly FolioAIEditNormalization[];
  from: number;
  to: number;
  /**
   * A deletion of the whole paragraph by its identity, with a block after
   * it: an insertion beside that paragraph keeps its own anchor.
   */
  removesParagraphBeforeBlock?: true;
  /**
   * Resolved beside a paragraph no paragraph follows, counting on the
   * batch's other half there: the deletion of that paragraph, or an
   * insertion after it. Without that half it resolves again on its own.
   */
  besideDeletion?: true;
} & (
  | { type: "plain"; operation: FolioDocumentOperation }
  | {
      type: "commented";
      operation: CommentOperation;
      comment: {
        text: string;
        from: TextPosition;
        to: TextPosition;
        /**
         * The comment covers the paragraphs the operation inserts, whose
         * identities the compiler gives: its range is read after compiling.
         */
        inserted?: true;
      };
    }
);

type ResolveCommentOptions = {
  operation: FolioDocumentOperation;
  from: TextPosition;
  to: TextPosition;
};

const resolveComment = ({ operation, from, to }: ResolveCommentOptions) => {
  if ("comment" in operation && operation.comment !== undefined) {
    return {
      type: "commented",
      operation,
      comment: { text: operation.comment.text, from, to },
    } as const satisfies Pick<
      Extract<ResolvedOperation, { type: "commented" }>,
      "type" | "operation" | "comment"
    >;
  }
  return { type: "plain", operation } as const;
};

const SPACING_FIELDS = [
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "beforeAutospacing",
  "afterAutospacing",
] as const satisfies readonly (keyof FolioAIParagraphSpacing)[];
const INDENTATION_FIELDS = [
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
] as const satisfies readonly (keyof FolioAIParagraphIndentation)[];

/**
 * The stated paragraph properties an operation writes: each given property
 * as it says, a spacing or indentation object as the whole direct child, so
 * an attribute it leaves out is removed, and `null` clears a property back
 * to what the style gives.
 */
const statedParagraphPatch = (
  properties: FolioAIBlockParagraphProperties,
): ParagraphPropsPatch => {
  const patch: Record<string, unknown> = {};
  if (properties.styleId !== undefined) patch["styleId"] = properties.styleId;
  if (properties.outlineLevel !== undefined)
    patch["outlineLevel"] = properties.outlineLevel;
  if (properties.alignment !== undefined)
    patch["alignment"] = properties.alignment;
  // Stated numbering as given; inheriting removes the stated value.
  const { numbering } = properties;
  if (numbering !== undefined && numbering.kind !== "newList")
    patch["numPr"] = numbering.kind === "inherit" ? null : numbering;
  const { spacing, indentation } = properties;
  if (spacing !== undefined)
    for (const field of SPACING_FIELDS)
      patch[field] = spacing === null ? null : (spacing[field] ?? null);
  if (indentation !== undefined)
    for (const field of INDENTATION_FIELDS)
      patch[field] = indentation === null ? null : (indentation[field] ?? null);
  // Every key is a ParagraphFormatting field set to its value or null.
  return patch as ParagraphPropsPatch;
};

type ResolveParagraphPropertiesOptions = {
  mintList: (format: FolioAIListKind) => MintedList;
  operation: Extract<
    FolioDocumentOperation,
    { type: "setBlockParagraphProperties" }
  >;
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>;
  projection: CanonicalProjection;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

const resolveParagraphProperties = ({
  operation,
  paragraph,
  projection,
  mintList,
  refusal,
}: ResolveParagraphPropertiesOptions) => {
  const requested = operation.properties.numbering;
  const stated = statedNumbering(requested, mintList);
  const patch = statedParagraphPatch({
    ...operation.properties,
    ...(stated.numbering === undefined ? {} : { numbering: stated.numbering }),
  });
  if (Object.keys(patch).length === 0) return refusal("emptyOperation");
  const at = projection.addressAt(paragraph.start);
  if (at.isErr()) return refusal("unsupportedBlock");
  const intents: EditorIntent[] = [
    {
      type: "formatParagraph",
      at: at.value,
      patch,
      ...(stated.newList === undefined ? {} : { newList: stated.newList }),
    },
  ];
  return Result.ok({
    type: "plain" as const,
    operation,
    intents,
    normalizations: [],
    from: paragraph.start,
    to: paragraph.start + paragraph.node.content.size,
  });
};

/**
 * The intent that states `properties` on one paragraph by its identity,
 * none when they state nothing.
 */
const statedPropertiesIntents = ({
  story,
  blockId,
  properties,
  mintList,
}: {
  story: OpStory;
  blockId: string;
  properties: FolioAIBlockParagraphProperties | undefined;
  mintList: (format: FolioAIListKind) => MintedList;
}): EditorIntent[] => {
  if (properties === undefined) return [];
  const stated = statedNumbering(properties.numbering, mintList);
  const patch = statedParagraphPatch({
    ...properties,
    ...(stated.numbering === undefined ? {} : { numbering: stated.numbering }),
  });
  if (Object.keys(patch).length === 0) return [];
  return [
    {
      type: "formatParagraph",
      at: { story, blockId, offset: 0 },
      patch,
      ...(stated.newList === undefined ? {} : { newList: stated.newList }),
    },
  ];
};

type BlockInsertion = Extract<
  FolioDocumentOperation,
  { type: "insertAfterBlock" | "insertBeforeBlock" }
>;

const isBlockInsertion = (
  operation: FolioDocumentOperation,
): operation is BlockInsertion =>
  operation.type === "insertAfterBlock" ||
  operation.type === "insertBeforeBlock";

type TableInsertion = Extract<FolioDocumentOperation, { type: "insertTable" }>;

/** An operation that puts a new block, a paragraph or a table, beside a paragraph. */
const placesBlock = (
  operation: FolioDocumentOperation,
): operation is BlockInsertion | TableInsertion =>
  isBlockInsertion(operation) || operation.type === "insertTable";

type ProjectedParagraph = NonNullable<
  ReturnType<CanonicalProjection["paragraph"]>
>;

type ResolveBlockInsertionOptions = {
  mintList: (format: FolioAIListKind) => MintedList;
  operation: BlockInsertion;
  /** The paragraph the insertion goes beside. */
  paragraph: ProjectedParagraph;
  /** The paragraph whose stated properties the new paragraph copies. */
  formattingSource: ProjectedParagraph;
  projection: CanonicalProjection;
  story: OpStory;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

/**
 * Where an insertion beside `paragraph` goes: beside the paragraph, or, for a
 * paragraph in a table cell, beside the outermost table, as the default
 * engine places it. The point names the paragraph on the far side of the
 * table; none stands there when the table starts or ends its container.
 */
const blockInsertionPlace = (
  operation: BlockInsertion,
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>,
  projection: CanonicalProjection,
): { at: BlockInsertionPoint; position: number } | undefined => {
  const side = operation.type === "insertAfterBlock" ? "after" : "before";
  const table = findOutermostTableBoundary(projection.doc, paragraph.start - 1);
  if (table === null)
    return {
      at: { type: side, blockId: paragraph.blockId },
      position:
        side === "after"
          ? paragraph.start + paragraph.node.content.size
          : paragraph.start,
    };
  const position = table[side];
  const resolved = projection.doc.resolve(position);
  const beside = side === "after" ? resolved.nodeAfter : resolved.nodeBefore;
  if (beside?.type.name !== "paragraph") return undefined;
  const neighbour = projection.paragraphStartingAt(
    side === "after" ? position : position - beside.nodeSize,
  );
  if (neighbour === undefined) return undefined;
  // After the table is before the paragraph that follows it, and before it
  // is after the one it follows.
  return {
    at: {
      type: side === "after" ? "before" : "after",
      blockId: neighbour.blockId,
    },
    position,
  };
};

/** Where an insertion beside a paragraph goes, and whose properties it copies. */
type BesidePendingJoin = {
  place: ProjectedParagraph;
  formattingSource: ProjectedParagraph;
};

/** A suggestion reads the document as it stands: beside the paragraph it names. */
const besideAsItStands = (
  paragraph: ProjectedParagraph,
): BesidePendingJoin => ({
  place: paragraph,
  formattingSource: paragraph,
});

type BesidePendingJoinOptions = {
  side: "before" | "after";
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  noteReferences: NoteReferenceLabels;
};

/**
 * Where an insertion beside `paragraph` goes, read as the joined paragraph a
 * pending join makes of it in the document the batch read
 * (`pendingJoinParagraphs`, which the default engine places by too): after
 * the join's last paragraph, or before its first, copying the properties of
 * the paragraph the join leaves, which the joined paragraph has. The batch's
 * other operations do not move it: a split of the paragraph the join leaves
 * puts the block after both halves, as one put after that paragraph goes.
 * Undefined beside a paragraph whose own break is pending removal where no
 * join reads it: after it, when the join leaves no paragraph with words for
 * the block to follow; before it, when the paragraph goes whole.
 */
const besidePendingJoin = ({
  side,
  paragraph,
  projection,
  noteReferences,
}: BesidePendingJoinOptions): BesidePendingJoin | undefined => {
  const options = {
    doc: projection.doc,
    paragraphPos: paragraph.start - 1,
    noteReferences,
  };
  const join = pendingJoinParagraphs(options);
  const breakGoes = isDeletedPPrMark(paragraph.node.attrs["pPrMark"]);
  const after = side === "after";
  if (
    breakGoes &&
    (after ? join.last === options.paragraphPos : !joinsIn(options))
  )
    return undefined;
  const at = (position: number) =>
    position === options.paragraphPos
      ? paragraph
      : projection.paragraphStartingAt(position);
  const place = at(after ? join.last : join.first);
  const formattingSource = at(join.last);
  if (place === undefined || formattingSource === undefined)
    return panic("A pending join names a paragraph the projection lacks.");
  return { place, formattingSource };
};

/**
 * One new paragraph beside the anchor, as the default engine writes it: the
 * anchor's stated paragraph properties unless `inheritFormatting` is false,
 * with the operation's own on top, holding the text as one plain run.
 */
const resolveBlockInsertion = ({
  operation,
  paragraph,
  formattingSource,
  projection,
  story,
  mintList,
  refusal,
}: ResolveBlockInsertionOptions) => {
  const requested = operation.numbering;
  const stated = statedNumbering(requested, mintList);
  const { text } = operation;
  if (stripInlineEmphasisMarkers(text) !== text)
    return refusal(
      "unsupportedBlock",
      "Canonical public emphasis replacements are unavailable.",
    );
  const inline = operation.lineBreakMode === "inline";
  if (/\[\^(?:e)?\d+\]/u.test(text))
    return refusal(
      "unsupportedBlock",
      "Canonical block insertion requires plain text.",
    );
  // Each line its own paragraph, blank lines dropped, as the default engine
  // splits them; inline, line breaks stay in the one paragraph. A tab is a
  // tab element either way.
  const split = !inline && LINE_BREAK_PATTERN.test(text);
  const lines = split
    ? text.split(LINE_BREAK_PATTERN).filter((line) => line.trim().length > 0)
    : [text];
  const texts = lines.length === 0 ? [""] : lines;
  const formatting: Record<string, unknown> =
    operation.inheritFormatting === false
      ? {}
      : { ...formattingSource.source.formatting };
  for (const [key, value] of Object.entries(
    statedParagraphPatch({
      ...(operation.styleId === undefined
        ? {}
        : { styleId: operation.styleId }),
      ...(operation.outlineLevel === undefined
        ? {}
        : { outlineLevel: operation.outlineLevel }),
      ...(operation.alignment === undefined
        ? {}
        : { alignment: operation.alignment }),
      ...(operation.spacing === undefined
        ? {}
        : { spacing: operation.spacing }),
      ...(operation.indentation === undefined
        ? {}
        : { indentation: operation.indentation }),
      ...(stated.numbering === undefined
        ? {}
        : { numbering: stated.numbering }),
    }),
  ))
    if (value === null) Reflect.deleteProperty(formatting, key);
    else formatting[key] = value;
  const everyParagraph = operation.formattingScope === "allParagraphs";
  const paragraphs = texts.map((line, index): Paragraph => {
    // The first paragraph takes the operation's formatting; later ones are
    // body paragraphs unless the scope formats every paragraph alike.
    const own: Record<string, unknown> =
      index === 0 || everyParagraph ? { ...formatting } : {};
    if (index === 0 && operation.pageBreakBefore === true)
      own["pageBreakBefore"] = true;
    // An authored page break stands in for the first paragraph's text.
    const pageBreak = index === 0 ? operation.hardPageBreak : undefined;
    let content: Paragraph["content"] = [];
    if (pageBreak !== undefined)
      content = [
        {
          type: "run",
          content: [
            {
              type: "break",
              breakType: "page",
              ...(pageBreak.clear === undefined
                ? {}
                : { clear: pageBreak.clear }),
            },
          ],
        },
      ];
    else if (line !== "")
      content = [{ type: "run", content: inlineText(line) }];
    return {
      type: "paragraph",
      content,
      // The anchor's properties with the operation's set or cleared over them.
      ...(Object.keys(own).length === 0
        ? {}
        : { formatting: own as ParagraphFormatting }),
    };
  });
  const place = blockInsertionPlace(operation, paragraph, projection);
  if (place === undefined)
    return refusal(
      "unsupportedBlock",
      "Canonical insertion beside a table needs a paragraph on that side of it.",
    );
  // The anchor keeps its identity; the new paragraph takes a fresh one.
  const intents: EditorIntent[] = [
    {
      type: "insertParagraphs",
      story,
      at: place.at,
      paragraphs,
      ...(stated.newList === undefined ? {} : { newList: stated.newList }),
    },
  ];
  const { position } = place;
  // A comment covers the inserted text, so an insertion with no text has
  // nothing for it to cover, as the default engine finds.
  if (
    operation.comment !== undefined &&
    texts.every((line) => line.length === 0)
  )
    return refusal("unsupportedBlock");
  const normalizations: FolioAIEditNormalization[] = split
    ? [
        {
          id: operation.id,
          code: "splitMultilineText",
          paragraphCount: texts.length,
        },
      ]
    : [];
  const resolved = { intents, normalizations, from: position, to: position };
  if (operation.comment === undefined)
    return Result.ok({ type: "plain" as const, operation, ...resolved });
  // Where the inserted paragraphs are is known once they are compiled.
  const anchor = { story, blockId: place.at.blockId, offset: 0 };
  return Result.ok({
    type: "commented" as const,
    operation,
    comment: {
      text: operation.comment.text,
      from: anchor,
      to: anchor,
      inserted: true as const,
    },
    ...resolved,
  });
};

const LINE_BREAK_PATTERN = /\r\n|\r|\n/u;

type ResolveTableInsertionOptions = {
  operation: TableInsertion;
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  story: OpStory;
  noteReferences: NoteReferenceLabels;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  tableGap: (
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

/**
 * A table beside the anchor, built from the cell texts as the default engine
 * builds it (`buildTableNode`) and placed as a paragraph beside the anchor is
 * (`besidePendingJoin`): between the paragraph it goes beside and the one
 * next to it, and after a container's last paragraph with the empty
 * paragraph the container ends with. Beside a table, the anchor's own or a
 * neighbouring one, or any block other than a paragraph, it stays under the
 * table structure gap.
 */
const resolveTableInsertion = ({
  operation,
  paragraph,
  projection,
  story,
  noteReferences,
  refusal,
  tableGap,
}: ResolveTableInsertionOptions) => {
  const side = operation.position ?? "after";
  if (findOutermostTableBoundary(projection.doc, paragraph.start - 1) !== null)
    return tableGap("Canonical table insertion beside a table is unavailable.");
  const placed = besidePendingJoin({
    side,
    paragraph,
    projection,
    noteReferences,
  });
  if (placed === undefined) return refusal("pendingDeletion");
  const { place } = placed;
  const $place = projection.doc.resolve(place.start - 1);
  const neighbour = $place.parent.maybeChild(
    $place.index() + (side === "after" ? 1 : -1),
  );
  if (neighbour !== null && neighbour.type.name !== "paragraph")
    return tableGap(
      "Canonical table insertion needs a paragraph or the container's edge beside it.",
    );
  const node = buildTableNode({
    schema: projection.doc.type.schema,
    rows: operation.rows,
  });
  if (node === null) return refusal("unsupportedBlock");
  const attrs = expectTableAttrs(node);
  const formatting = tableAttrsToFormatting(attrs);
  const table: Table = {
    type: "table",
    ...(formatting === undefined ? {} : { formatting }),
    ...(attrs.columnWidths === undefined
      ? {}
      : { columnWidths: [...attrs.columnWidths] }),
    rows: standaloneTableRowsFromProseMirror(node),
  };
  const position =
    side === "after" ? place.start + place.node.content.size : place.start;
  const intent: EditorIntent = {
    type: "tableBlock",
    operation: {
      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
      story,
      blockId: place.blockId,
      table,
      side,
      // After a container's last paragraph the container still ends with one.
      paragraphAfter: side === "after" && neighbour === null ? "empty" : "none",
    },
  };
  return Result.ok({
    type: "plain" as const,
    operation,
    intents: [intent],
    normalizations: [],
    from: position,
    to: position,
  });
};

/**
 * The operations a canonical suggestion can carry: the default engine's;
 * splits and merges, which the canonical session records as tracked paragraph
 * marks stamped with the suggestion; and comments, which it holds as proposed
 * by the suggestion until it is accepted or rejected.
 */
const CANONICAL_SUGGESTED_OPERATION_TYPES: ReadonlySet<
  FolioDocumentOperation["type"]
> = new Set<FolioDocumentOperation["type"]>([
  ...SUGGESTED_SUPPORTED_OPERATION_TYPES,
  "splitBlock",
  "mergeBlockWithNext",
  "commentOnRange",
  "commentOnBlock",
]);

/**
 * Whether one of two resolved operations inserts beside the paragraph the
 * other deletes by identity: they touch only at that paragraph's edge, and
 * both apply, as in the default engine. A deletion of a paragraph that ends
 * its container still refuses the insertion beside it.
 */
const besideItsDeletion = (
  left: Pick<ResolvedOperation, "operation" | "removesParagraphBeforeBlock">,
  right: Pick<ResolvedOperation, "operation" | "removesParagraphBeforeBlock">,
): boolean => {
  const pair = (
    { operation: insertion }: typeof left,
    { operation: deletion, removesParagraphBeforeBlock }: typeof left,
  ) =>
    (insertion.type === "insertAfterBlock" ||
      insertion.type === "insertBeforeBlock") &&
    deletion.type === "deleteBlock" &&
    removesParagraphBeforeBlock === true &&
    insertion.blockId === deletion.blockId;
  return pair(left, right) || pair(right, left);
};

/**
 * Whether one of two resolved operations inserts a block right before or
 * right after the paragraph the other splits: the split claims that
 * paragraph whole, and the insertion touches only its edge. Both apply, as
 * in the default engine: the insertion goes in first, beside the paragraph
 * it read, and the split then divides that paragraph by its identity.
 */
const besideItsSplit = (
  left: Pick<ResolvedOperation, "operation" | "from" | "to">,
  right: Pick<ResolvedOperation, "operation" | "from" | "to">,
): boolean => {
  const pair = (insertion: typeof left, split: typeof left) =>
    placesBlock(insertion.operation) &&
    insertion.from === insertion.to &&
    split.operation.type === "splitBlock" &&
    (insertion.from === split.from || insertion.from === split.to);
  return pair(left, right) || pair(right, left);
};

/** A block operation without the move it names: a plain insertion or deletion. */
const withoutMove = (
  operation: FolioDocumentOperation,
): FolioDocumentOperation => {
  if (
    operation.type !== "deleteBlock" &&
    operation.type !== "insertAfterBlock" &&
    operation.type !== "insertBeforeBlock"
  )
    return operation;
  const { moveId: _moveId, ...plain } = operation;
  return plain;
};

/** The move a block operation names, if it is a half of one. */
const blockMoveId = (operation: FolioDocumentOperation): string | undefined =>
  operation.type === "deleteBlock" ||
  operation.type === "insertAfterBlock" ||
  operation.type === "insertBeforeBlock"
    ? operation.moveId
    : undefined;

/**
 * A line's text with each tab and line break as the element it is; CRLF, CR
 * and LF are each one line break.
 */
const inlineText = (line: string): Run["content"] =>
  line.split(/(\t|\r\n|\r|\n)/u).flatMap((part): Run["content"] => {
    if (part === "\t") return [{ type: "tab" }];
    if (LINE_BREAK_PATTERN.test(part)) return [{ type: "break" }];
    return part === "" ? [] : [{ type: "text", text: part }];
  });

type ResolveBlockDeletionOptions = {
  operation: Extract<FolioDocumentOperation, { type: "deleteBlock" }>;
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>;
  projection: CanonicalProjection;
  projectionStory: OpStory;
  doc: PMNode;
  tracked: boolean;
  /**
   * Whether the batch inserts a paragraph beside this one, which no
   * paragraph follows: the insertion applies first, and the deletion then
   * finds this paragraph by identity.
   */
  gainsFollower: boolean;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

type CanonicalPublicOperationsOptions = {
  session: CanonicalSession;
  getState: (story: OpStory) => EditorState | null;
  publish: (commit: CanonicalCommit) => boolean;
};

type CompilePublicIntentsOptions = {
  document: Document;
  intents: readonly EditorIntent[];
  mode: "direct" | "tracked-changes";
  author: string;
  date: string;
  /** The pending suggestion the records belong to, in suggested mode. */
  suggestionId?: string | undefined;
  nextRevisionId: number;
  allocate: ReturnType<typeof createEditorIntentIdAllocator>;
  /** Paragraph identities the session has consumed, retired ones included. */
  reservedBlockIds: ReadonlySet<string>;
  /**
   * The ids the session may mint, when it holds a lease: revision ids then
   * come from the allocator's leased ones rather than one contiguous range.
   */
  idLease: IdLease | undefined;
};

const compilePublicIntents = ({
  document,
  intents,
  mode,
  author,
  date,
  suggestionId,
  nextRevisionId,
  allocate,
  reservedBlockIds,
  idLease,
}: CompilePublicIntentsOptions) => {
  const ops: DocumentOp[] = [];
  const revisions: number[] = [];
  // A split's new half takes the id allocated here, not the one it resolved
  // with: a later intent naming that half names the allocated one.
  const allocated = new Map<string, string>();
  for (const sourceIntent of intents) {
    const ids = allocate(document, sourceIntent, idLease);
    const leased = idLease === undefined ? {} : { idLease };
    let intent: EditorIntent = sourceIntent;
    if (sourceIntent.type === "splitParagraph") {
      allocated.set(sourceIntent.newBlockId, ids.newBlockId);
      intent = {
        type: "splitParagraph",
        at: sourceIntent.at,
        ...(sourceIntent.to === undefined ? {} : { to: sourceIntent.to }),
        ...(sourceIntent.newHalf === undefined
          ? {}
          : { newHalf: sourceIntent.newHalf }),
        newBlockId: ids.newBlockId,
      };
    } else if (sourceIntent.type === "formatParagraph") {
      const blockId = allocated.get(sourceIntent.at.blockId);
      if (blockId !== undefined)
        intent = { ...sourceIntent, at: { ...sourceIntent.at, blockId } };
    }
    const firstRevisionId = nextRevisionId;
    const compiled = compileEditorIntent(document, {
      intent,
      mode:
        mode === "direct"
          ? {
              type: "editing",
              newIds: ids.newIds,
              reservedBlockIds,
              ...leased,
            }
          : {
              type: "suggesting",
              reservedBlockIds,
              ...leased,
              revision: {
                id: idLease === undefined ? firstRevisionId : ids.revisionId,
                author,
                date,
                ...(suggestionId === undefined ? {} : { suggestionId }),
              },
              newIds: {
                control: ids.newIds.control,
                revision:
                  idLease === undefined
                    ? ids.newIds.revision.map(
                        (_, index) => firstRevisionId + index + 1,
                      )
                    : ids.newIds.revision,
              },
            },
    });
    if (compiled.isErr()) return Result.err(compiled.error);
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    document = applied.value.document;
    ops.push(...compiled.value.ops);
    revisions.push(...applied.value.revisions);
    for (const id of applied.value.revisions)
      nextRevisionId = Math.max(nextRevisionId, id + 1);
  }
  return Result.ok({ document, ops, revisions, nextRevisionId });
};

const CanonicalPublicOperationResolutionErrorBase: TaggedErrorClass<"CanonicalPublicOperationResolutionError"> =
  TaggedError("CanonicalPublicOperationResolutionError");

class CanonicalPublicOperationResolutionError extends CanonicalPublicOperationResolutionErrorBase<{
  message: string;
  reason: FolioAIEditSkipReason;
  /** The capability the refusal waits for, when it is not the compiler's inline one. */
  gap?: CanonicalGap;
}> {}

type ResolveCanonicalPublicOperationOptions = {
  operation: FolioDocumentOperation;
  options: CanonicalPublicOperationOptions;
  state: EditorState;
  story: OpStory;
  diff: ReturnType<typeof wordDiffSessionFromOptions>["diff"];
  /** Whether the batch records its edits as revisions. */
  tracked: boolean;
  /**
   * Whether the batch records its edits as suggestions, which read the
   * document as it stands rather than as accepting leaves it.
   */
  suggested: boolean;
  /** A new list of one kind for this operation, minted after the batch's earlier ones. */
  mintList: (format: FolioAIListKind) => MintedList;
  /** The blocks the batch deletes, and those it inserts beside. */
  besideDeletions: BesideDeletions;
};

type BesideDeletions = {
  deleted: ReadonlySet<string>;
  insertedBeside: ReadonlySet<string>;
};

/**
 * The stated numbering a request names, minting its new list once for the
 * operation, as the default engine allocates one instance per operation.
 */
const statedNumbering = (
  numbering: FolioAIListNumbering | undefined,
  mint: (format: FolioAIListKind) => MintedList,
): { numbering?: FolioContentStatedNumbering; newList?: NewListInstance } => {
  if (numbering === undefined) return {};
  if (numbering.kind !== "newList") return { numbering };
  const minted = mint(numbering.format);
  return {
    numbering: paragraphNumberingReference({
      numId: minted.numId,
      ilvl: numbering.level ?? 0,
    }),
    newList: minted.newList,
  };
};

const canonicalPublicStory = (story: FolioDocumentOperationStory): OpStory => {
  if (story === "main") return story;
  switch (story.type) {
    case "header":
    case "footer":
      return { kind: story.type, rId: story.relationshipId };
    case "footnote":
    case "endnote":
      return { kind: story.type, id: story.noteId };
    default: {
      const unreachable: never = story;
      return panic(`Unknown public operation story ${unreachable}`);
    }
  }
};

const batchStatus = (
  batch: FolioDocumentOperationBatch,
  skippedCount: number,
) => {
  if (batch.dryRun) return "previewed" as const;
  if (batch.atomic && skippedCount > 0) return "rejected" as const;
  return "committed" as const;
};

/** Handles address isolated journal groups; no second document or PM history is retained. */
export class CanonicalPublicOperations {
  private readonly options: CanonicalPublicOperationsOptions;
  private readonly handles: {
    handle: FolioDocumentOperationUndoHandle;
    version: number;
    story: OpStory;
  }[] = [];
  private nextHandle = 1;
  private readonly handleNamespace = crypto.randomUUID();

  constructor(options: CanonicalPublicOperationsOptions) {
    this.options = options;
  }

  apply(
    options: CanonicalPublicOperationOptions,
  ): FolioDocumentOperationResult {
    const { session } = this.options;
    const parsed = parseFolioDocumentOperationBatch(options.batch);
    // A suggested operation never folds into a suggestion already pending,
    // here or in the host's other stories.
    const batch =
      parsed.mode !== "suggested" ||
      options.restagesSuggestions === true ||
      session.isComposing
        ? parsed
        : withDistinctSuggestionIds(
            parsed,
            new Set([
              ...pendingSuggestionIdsIn(readModel(session.document)),
              ...(options.pendingSuggestionIds ?? []),
            ]),
          );
    const publicStory = options.story ?? "main";
    const story = canonicalPublicStory(publicStory);
    const state = this.options.getState(story);
    const projected = session.projectStory(story);
    const skipped: FolioAIEditSkippedOperation[] = [];
    const normalizations: FolioAIEditNormalization[] = [];
    const refusals = new Map<string, CanonicalPublicOperationRefusal>();
    const resolved: ResolvedOperation[] = [];
    const mintList = createBatchListMinter(() => readModel(session.document));
    const { diff } = wordDiffSessionFromOptions(options.wordDiff);
    let nextRevisionId = options.revisionStamp?.idSeed ?? 1;
    if (!session.isComposing && options.revisionStamp === undefined) {
      for (const currentStory of documentStories(readModel(session.document))) {
        const currentProjection = session.projectStory(currentStory);
        if (currentProjection.isOk())
          nextRevisionId = Math.max(
            nextRevisionId,
            maxAnnotationIdInDoc(currentProjection.value.doc) + 1,
          );
      }
    }
    const skip = (
      id: string,
      reason: FolioAIEditSkipReason,
      message?: string,
    ) => {
      skipped.push({
        id,
        reason,
        ...(message === undefined ? {} : { message }),
      });
    };
    const result = (
      applied: FolioDocumentOperationResult["applied"],
      undoHandle: FolioDocumentOperationUndoHandle | null,
    ): FolioDocumentOperationResult => {
      const failures = new Map(skipped.map((item) => [item.id, item]));
      const ordered = batch.operations.flatMap(({ id }) => {
        const failure = failures.get(id);
        if (!failure) return [];
        const canonicalRefusal = refusals.get(id);
        return [canonicalRefusal ? { ...failure, canonicalRefusal } : failure];
      });
      return {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        status: batchStatus(batch, ordered.length),
        applied,
        skipped: ordered,
        issues: getFolioDocumentOperationIssues(batch.operations, ordered),
        receipts: getFolioDocumentOperationReceiptsForStory({
          operations: batch.operations,
          applied,
          story: publicStory,
        }),
        undoHandle,
        nextRevisionId,
        ...(normalizations.length === 0 ? {} : { normalizations }),
      };
    };
    const besideDeletions: BesideDeletions = {
      deleted: new Set(
        batch.operations.flatMap((operation) =>
          operation.type === "deleteBlock" ? [operation.blockId] : [],
        ),
      ),
      insertedBeside: new Set(
        batch.operations.flatMap((operation) =>
          operation.type === "insertAfterBlock" ||
          operation.type === "insertBeforeBlock"
            ? [operation.blockId]
            : [],
        ),
      ),
    };
    const overlapsResolved = (candidate: ResolvedOperation) =>
      resolved.some(
        (earlier) =>
          earlier !== candidate &&
          !besideItsDeletion(earlier, candidate) &&
          !besideItsSplit(earlier, candidate) &&
          ((candidate.from < earlier.to && candidate.to > earlier.from) ||
            (earlier.from === earlier.to &&
              candidate.from <= earlier.from &&
              candidate.to >= earlier.from) ||
            (candidate.from === candidate.to &&
              candidate.from >= earlier.from &&
              candidate.from <= earlier.to)),
      );
    for (const operation of batch.operations) {
      if (!state || projected.isErr() || session.isComposing) {
        skip(operation.id, "documentNotEditable");
        continue;
      }
      if (
        batch.mode === "suggested" &&
        !CANONICAL_SUGGESTED_OPERATION_TYPES.has(operation.type)
      ) {
        skip(operation.id, "unsupportedMode");
        continue;
      }
      if (
        options.tableTemplates !== undefined ||
        options.undefinedReferences === "keep"
      ) {
        refusals.set(operation.id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(
          operation.id,
          "unsupportedBlock",
          "This canonical compiler does not support the supplied execution options.",
        );
        continue;
      }
      const disposition =
        CANONICAL_PUBLIC_OPERATION_DISPOSITIONS[operation.type];
      if (disposition !== "compile") {
        refusals.set(operation.id, {
          gap: disposition,
        });
        skip(operation.id, "unsupportedBlock");
        continue;
      }
      const compiled = this.resolve({
        operation,
        options,
        state,
        story,
        diff,
        tracked: batch.mode !== "direct",
        suggested: batch.mode === "suggested",
        mintList,
        besideDeletions,
      });
      if (compiled.isErr()) {
        const gap =
          compiled.error.gap ??
          (compiled.error.reason === "unsupportedBlock"
            ? CANONICAL_GAP.publicUnsupportedInline
            : undefined);
        if (gap !== undefined) refusals.set(operation.id, { gap });
        skip(operation.id, compiled.error.reason, compiled.error.message);
        continue;
      }
      if (overlapsResolved(compiled.value)) {
        skip(operation.id, "overlappingOperation");
        continue;
      }
      resolved.push(compiled.value);
    }
    // An operation resolved beside a deleted paragraph that ends its run
    // counts on its other half; where that half did not resolve, it resolves
    // again on its own, against the operations that did.
    for (const item of [...resolved]) {
      if (item.besideDeletion !== true || !state) continue;
      const { operation } = item;
      if (
        operation.type !== "deleteBlock" &&
        operation.type !== "insertAfterBlock" &&
        operation.type !== "insertBeforeBlock"
      )
        continue;
      const partnered = resolved.some(({ operation: other }) =>
        operation.type === "deleteBlock"
          ? (other.type === "insertAfterBlock" ||
              other.type === "insertBeforeBlock") &&
            other.blockId === operation.blockId
          : other.type === "deleteBlock" && other.blockId === operation.blockId,
      );
      if (partnered) continue;
      resolved.splice(resolved.indexOf(item), 1);
      const alone = this.resolve({
        operation,
        options,
        state,
        story,
        diff,
        tracked: batch.mode !== "direct",
        suggested: batch.mode === "suggested",
        mintList,
        besideDeletions: {
          deleted: new Set(),
          insertedBeside: new Set(),
        },
      });
      if (alone.isErr()) {
        skip(operation.id, alone.error.reason, alone.error.message);
        continue;
      }
      if (overlapsResolved(alone.value)) {
        skip(operation.id, "overlappingOperation");
        continue;
      }
      resolved.push(alone.value);
    }
    // A move id names one relocation, a move only when it reaches exactly one
    // source and one destination, as the default engine pairs them; an
    // unpaired id records a plain insertion or deletion and is reported.
    const moveSides = new Map<string, { from: number; to: number }>();
    for (const { operation } of resolved) {
      const moveId = blockMoveId(operation);
      if (moveId === undefined) continue;
      const sides = moveSides.get(moveId) ?? { from: 0, to: 0 };
      if (operation.type === "deleteBlock") sides.from += 1;
      else sides.to += 1;
      moveSides.set(moveId, sides);
    }
    // A suggested destination before its own source would cut the move.
    const unpairable = new Set(
      batch.mode === "suggested"
        ? resolved.flatMap(({ operation }) =>
            operation.type === "insertBeforeBlock" &&
            operation.moveId !== undefined &&
            resolved.some(
              (candidate) =>
                candidate.operation.type === "deleteBlock" &&
                candidate.operation.moveId === operation.moveId &&
                candidate.operation.blockId === operation.blockId,
            )
              ? [operation.moveId]
              : [],
          )
        : [],
    );
    const paired = (moveId: string) =>
      !unpairable.has(moveId) &&
      moveSides.get(moveId)?.from === 1 &&
      moveSides.get(moveId)?.to === 1;
    const pairedMoves = new Set<string>();
    for (const [index, item] of resolved.entries()) {
      const moveId = blockMoveId(item.operation);
      if (moveId === undefined) continue;
      if (!paired(moveId)) {
        normalizations.push({
          id: item.operation.id,
          code: "unpairedMove",
          moveId,
        });
        continue;
      }
      pairedMoves.add(moveId);
      const intents: EditorIntent[] = [];
      for (const intent of item.intents)
        intents.push(
          intent.type === "insertParagraphs" ||
            intent.type === "deleteParagraphs"
            ? Object.assign({}, intent, { move: true as const })
            : intent,
        );
      resolved[index] = { ...item, intents };
    }
    if (batch.atomic && skipped.length > 0) {
      for (const { operation } of resolved)
        skip(operation.id, "atomicBatchRejected");
      return result([], null);
    }
    if (!state || projected.isErr() || session.isComposing)
      return result([], null);
    // Descending read positions keep independent snapshot coordinates valid;
    // at one position an insertion goes first, beside a paragraph the batch
    // then deletes.
    resolved.sort(
      (left, right) =>
        right.from - left.from || left.to - left.from - (right.to - right.from),
    );
    let document = readModel(session.document);
    const ops: DocumentOp[] = [];
    const appliedById = new Map<
      string,
      FolioDocumentOperationResult["applied"][number]
    >();
    const allocate = createEditorIntentIdAllocator({
      reservedBlockIds: session.reservedBlockIds,
    });
    const date = options.revisionStamp?.date ?? new Date().toISOString();
    for (const resolvedOperation of resolved) {
      const {
        operation,
        intents,
        normalizations: adjustments,
      } = resolvedOperation;
      // A suggested operation's records are one suggestion: its own id, or the
      // one it names.
      const suggestionId =
        batch.mode === "suggested"
          ? (("suggestionId" in operation
              ? operation.suggestionId
              : undefined) ?? operation.id)
          : undefined;
      const compiled = compilePublicIntents({
        document,
        intents,
        allocate,
        reservedBlockIds: session.reservedBlockIds,
        idLease: session.idLease,
        nextRevisionId,
        mode: batch.mode === "direct" ? "direct" : "tracked-changes",
        suggestionId,
        author: options.author ?? "AI",
        date,
      });
      if (compiled.isErr()) {
        refusals.set(operation.id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(operation.id, "unsupportedBlock", compiled.error.message);
        continue;
      }
      let nextDocument = compiled.value.document;
      const operationOps = [...compiled.value.ops];
      let commentId: number | undefined;
      switch (resolvedOperation.type) {
        case "plain":
          break;
        case "commented": {
          const { comment } = resolvedOperation;
          const id = freshCommentId(nextDocument, session.idLease);
          // A suggested operation's comment is proposed with it: pending,
          // never saved, accepted and rejected with the suggestion.
          const proposedBy = suggestionId === undefined ? {} : { suggestionId };
          if (comment.inserted === true) {
            // The comment covers every paragraph the operation inserted, at
            // any depth: in a content control or a table cell as much as in
            // the story's own block list.
            const paragraphsIn = (current: Document) => {
              const body = findStoryBody(current, story);
              return body === undefined
                ? []
                : storyParagraphs(body).map(({ paragraph }) => paragraph);
            };
            const known = new Set(
              paragraphsIn(document).flatMap(({ paraId }) =>
                paraId === undefined ? [] : [paraId],
              ),
            );
            const added = paragraphsIn(nextDocument).filter(
              ({ paraId }) => paraId !== undefined && !known.has(paraId),
            );
            const firstAdded = added.at(0);
            const lastAdded = added.at(-1);
            const commented =
              id.isErr() ||
              firstAdded?.paraId === undefined ||
              lastAdded?.paraId === undefined
                ? undefined
                : compileCanonicalComments({
                    document: nextDocument,
                    idLease: session.idLease,
                    command: {
                      type: "create",
                      comment: {
                        id: id.value,
                        author: options.author ?? "AI",
                        date,
                        done: false,
                        content: canonicalCommentBody(
                          nextDocument,
                          comment.text,
                          session.idLease,
                        ),
                        ...proposedBy,
                      },
                      anchor: {
                        kind: "range",
                        from: {
                          story,
                          blockId: firstAdded.paraId,
                          offset: 0,
                        },
                        to: {
                          story,
                          blockId: lastAdded.paraId,
                          offset: paragraphLength(lastAdded),
                        },
                      },
                    },
                  });
            if (commented === undefined || commented.isErr() || id.isErr()) {
              refusals.set(operation.id, {
                gap: CANONICAL_GAP.publicComments,
              });
              skip(
                operation.id,
                "unsupportedBlock",
                commented?.isErr() === true
                  ? commented.error.message
                  : "The inserted paragraphs cannot carry the comment.",
              );
              continue;
            }
            nextDocument = commented.value.document;
            operationOps.push(...commented.value.ops);
            commentId = id.value;
            break;
          }
          // The anchor paragraph at any depth: in a table cell or a content
          // control as much as in the story's own block list.
          const anchorIn = (current: Document) => {
            const body = findStoryBody(current, story);
            return body === undefined
              ? undefined
              : storyParagraphs(body).find(
                  ({ paragraph }) => paragraph.paraId === comment.from.blockId,
                )?.paragraph;
          };
          const before = anchorIn(document);
          const after = anchorIn(nextDocument);
          if (id.isErr() || before === undefined || after === undefined) {
            refusals.set(operation.id, { gap: CANONICAL_GAP.publicComments });
            skip(
              operation.id,
              "unsupportedBlock",
              id.isErr()
                ? id.error.message
                : "The comment anchor no longer exists.",
            );
            continue;
          }
          // A tracked edit's comment covers what the edit changed, from its
          // first revision to its last, as the default engine anchors it.
          const changed =
            batch.mode === "direct"
              ? undefined
              : revisionSpan(after.content, new Set(compiled.value.revisions));
          const from =
            changed === undefined
              ? comment.from
              : {
                  story: comment.from.story,
                  blockId: comment.from.blockId,
                  offset: changed.from,
                };
          const to =
            changed === undefined
              ? {
                  ...comment.to,
                  offset:
                    comment.to.offset +
                    paragraphLength(after) -
                    paragraphLength(before),
                }
              : {
                  story: comment.from.story,
                  blockId: comment.from.blockId,
                  offset: changed.to,
                };
          const commented = compileCanonicalComments({
            document: nextDocument,
            idLease: session.idLease,
            command: {
              type: "create",
              comment: {
                id: id.value,
                author: options.author ?? "AI",
                date,
                done: false,
                content: canonicalCommentBody(
                  nextDocument,
                  comment.text,
                  session.idLease,
                ),
                ...proposedBy,
              },
              anchor:
                to.offset === from.offset
                  ? { kind: "point", at: from }
                  : { kind: "range", from, to },
            },
          });
          if (commented.isErr()) {
            refusals.set(operation.id, { gap: CANONICAL_GAP.publicComments });
            skip(operation.id, "unsupportedBlock", commented.error.message);
            continue;
          }
          nextDocument = commented.value.document;
          operationOps.push(...commented.value.ops);
          commentId = id.value;
          break;
        }
        default: {
          const unreachable: never = resolvedOperation;
          panic(`Unknown resolved operation ${unreachable}`);
        }
      }
      if (operationOps.length === 0) {
        skip(operation.id, "noopOperation");
        continue;
      }
      // An operation a content-control lock refuses is skipped on its own,
      // as the document's rule; the rest of the batch still applies.
      const locked = lockViolationBetween(document, nextDocument);
      if (locked !== null) {
        skip(
          operation.id,
          CONTENT_CONTROL_LOCK,
          `Refused: ${lockedControlMessage(locked)}`,
        );
        continue;
      }
      normalizations.push(...adjustments);
      document = nextDocument;
      ops.push(...operationOps);
      nextRevisionId = Math.max(
        compiled.value.nextRevisionId,
        (commentId ?? -1) + 1,
      );
      const revisionIds = compiled.value.revisions;
      const revisionId = revisionIds.at(0);
      appliedById.set(operation.id, {
        id: operation.id,
        ...(commentId === undefined ? {} : { commentId }),
        ...(revisionId === undefined ? {} : { revisionId, revisionIds }),
        ...(suggestionId === undefined ? {} : { suggestionId }),
      });
    }
    // A move half that did not apply leaves the other a plain insertion or
    // deletion, reported unpaired: the batch applies again without the move.
    const applyingHalves = new Map<string, number>();
    for (const { operation } of resolved) {
      const moveId = blockMoveId(operation);
      if (moveId !== undefined && appliedById.has(operation.id))
        applyingHalves.set(moveId, (applyingHalves.get(moveId) ?? 0) + 1);
    }
    const orphaned = new Set(
      [...pairedMoves].filter((moveId) => applyingHalves.get(moveId) === 1),
    );
    if (orphaned.size > 0) {
      const orphanedMove = (operation: FolioDocumentOperation) => {
        const moveId = blockMoveId(operation);
        return moveId !== undefined && orphaned.has(moveId);
      };
      const plain = this.apply({
        ...options,
        batch: {
          ...batch,
          operations: batch.operations.map((operation) =>
            orphanedMove(operation) ? withoutMove(operation) : operation,
          ),
        },
      });
      const appliedIds = new Set(plain.applied.map(({ id }) => id));
      return {
        ...plain,
        normalizations: [
          ...(plain.normalizations ?? []),
          ...batch.operations.flatMap((operation) => {
            const moveId = blockMoveId(operation);
            return moveId !== undefined &&
              orphaned.has(moveId) &&
              appliedIds.has(operation.id)
              ? [{ id: operation.id, code: "unpairedMove" as const, moveId }]
              : [];
          }),
        ],
      };
    }
    if (batch.atomic && skipped.length > 0) {
      for (const { id } of appliedById.values())
        skip(id, "atomicBatchRejected");
      return result([], null);
    }
    if (ops.length === 0) return result([], null);
    const selection = projected.value.selectionAt(state);
    const prepared = selection.isErr()
      ? selection
      : session.prepareOps(state, ops, selection.value);
    if (prepared.isErr()) {
      for (const { id } of appliedById.values()) {
        refusals.set(id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(
          id,
          prepared.error.reason === "noChange"
            ? "noopOperation"
            : "unsupportedBlock",
          prepared.error.message,
        );
      }
      return result([], null);
    }
    const applied = batch.operations.flatMap(({ id }) => {
      const entry = appliedById.get(id);
      return entry ? [batch.dryRun ? { id } : entry] : [];
    });
    if (batch.dryRun) return result(applied, null);
    if (!this.options.publish(prepared.value)) {
      for (const { id } of applied) skip(id, "documentNotEditable");
      return result([], null);
    }
    const handle = {
      type: "documentOperationUndo",
      id: `canonical-${this.handleNamespace}-${String(this.nextHandle++)}`,
    } as const;
    this.handles.push({ handle, version: session.version, story });
    return result(applied, handle);
  }

  undo(
    handle: FolioDocumentOperationUndoHandle,
  ): FolioDocumentOperationUndoResult {
    const index = this.handles.findIndex(
      ({ handle: existing }) =>
        existing.id === handle.id && existing.type === handle.type,
    );
    if (index < 0)
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "unknownHandle",
      };
    if (index !== this.handles.length - 1)
      return { status: "rejected", undoHandle: handle, reason: "notLatest" };
    const entry = this.handles.at(-1);
    const { session } = this.options;
    if (!entry || entry.version !== session.version)
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "documentChanged",
      };
    const state = this.options.getState(entry.story);
    if (!state)
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "documentChanged",
      };
    const prepared = session.prepareUndo(state, entry.story);
    if (prepared.isErr() || !this.options.publish(prepared.value))
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "documentChanged",
      };
    this.handles.pop();
    // Undo advances the version while exposing the preceding unchanged journal group.
    const previous = this.handles.at(-1);
    if (previous && previous.version === entry.version - 1)
      previous.version = session.version;
    return { status: "undone", undoHandle: handle };
  }

  /**
   * A whole paragraph removed as the editor removes a selection of it with
   * its paragraph mark: from its start to the start of the paragraph after
   * it, or, for the paragraph that ends its block list, from the end of the
   * paragraph before it, whose mark then ends the list. Tracked, the text and
   * the mark are marked deleted.
   */
  private resolveBlockDeletion({
    operation,
    paragraph,
    projection,
    projectionStory,
    doc,
    tracked,
    gainsFollower,
    refusal,
  }: ResolveBlockDeletionOptions) {
    // A direct deletion leaves no text for a comment to cover, as the default
    // engine finds; a tracked one covers the deleted words.
    if (operation.comment !== undefined && !tracked)
      return refusal("unsupportedBlock");
    const end = paragraph.start + paragraph.node.content.size;
    const after = doc.resolve(end + 1).nodeAfter;
    const before = doc.resolve(paragraph.start - 1).nodeBefore;
    const neighbour = (node: PMNode | null) => {
      const id: unknown = node?.attrs["paraId"];
      return node?.type.name === "paragraph" && typeof id === "string"
        ? projection.paragraph(id)
        : undefined;
    };
    const next = neighbour(after);
    const previous = neighbour(before);
    // A tracked deletion's comment covers the deleted paragraph's words.
    const commentFrom = projection.addressAt(paragraph.start);
    const commentTo = projection.addressAt(end);
    if (commentFrom.isErr() || commentTo.isErr())
      return refusal("unsupportedBlock");
    const tracking = (
      resolved: Omit<
        Extract<ResolvedOperation, { type: "plain" }>,
        "type" | "operation"
      >,
    ): Result<ResolvedOperation, CanonicalPublicOperationResolutionError> =>
      Result.ok(
        operation.comment === undefined
          ? { type: "plain", operation, ...resolved }
          : {
              type: "commented",
              operation,
              comment: {
                text: operation.comment.text,
                from: commentFrom.value,
                to: commentTo.value,
              },
              ...resolved,
            },
      );
    // Before a table, tracked, the paragraph's own mark is the one deleted,
    // which a text range cannot say: the block deletion records it, as it
    // records a move's source, which the batch pairs before compiling.
    // So does a deletion the paragraph after it survives, by identity: an
    // insertion before that paragraph in the same batch leaves it in place.
    // With no paragraph after it, one the batch inserts beside it applies
    // first at one position: after it, that paragraph then follows it;
    // tracked, before it, the deletion finds it by identity there too, and
    // the inserted paragraph's own pending mark joins it. A move's halves do
    // the same, so no mark is left on the container's final paragraph.
    if (
      next !== undefined ||
      gainsFollower ||
      (tracked && (after !== null || operation.moveId !== undefined))
    ) {
      const intents: EditorIntent[] = [
        {
          type: "deleteParagraphs",
          story: projectionStory,
          blockIds: [paragraph.blockId],
        },
      ];
      return tracking({
        intents,
        normalizations: [],
        from: paragraph.start,
        to: end,
        ...(after === null && !gainsFollower
          ? {}
          : { removesParagraphBeforeBlock: true }),
      });
    }
    let range: { from: number; to: number } | undefined;
    // Removed directly, a paragraph no paragraph follows loses its words and
    // then joins the paragraph before it, which keeps its own identity and
    // properties; with none before it, it only loses its words.
    if (!tracked) {
      const own = (
        from: number,
        to: number,
      ): Result<EditorIntent[], CanonicalPublicOperationResolutionError> => {
        const start = projection.addressAt(from);
        const stop = projection.addressAt(to);
        return start.isErr() || stop.isErr()
          ? refusal("unsupportedBlock")
          : Result.ok(
              from === to
                ? []
                : [
                    {
                      type: "replaceText",
                      from: start.value,
                      to: stop.value,
                      text: "",
                    },
                  ],
            );
      };
      const cleared = own(paragraph.start, end);
      if (cleared.isErr()) return cleared;
      const intents: EditorIntent[] = [...cleared.value];
      if (previous !== undefined)
        intents.push({
          type: "joinParagraphs",
          story: projectionStory,
          blockId: previous.blockId,
          nextBlockId: paragraph.blockId,
          survivor: "first",
        });
      if (intents.length === 0) return refusal("noopOperation");
      return Result.ok({
        type: "plain" as const,
        operation,
        intents,
        normalizations: [],
        from: previous === undefined ? paragraph.start : previous.start,
        to: end,
      });
    } else if (previous !== undefined)
      range = {
        from: previous.start + previous.node.content.size,
        to: end,
      };
    if (range === undefined)
      return refusal(
        "unsupportedBlock",
        "Canonical deletion of the only paragraph in its container is unavailable.",
      );
    const from = projection.addressAt(range.from);
    const to = projection.addressAt(range.to);
    if (from.isErr() || to.isErr()) return refusal("unsupportedBlock");
    const intents: EditorIntent[] = [
      { type: "replaceText", from: from.value, to: to.value, text: "" },
    ];
    return tracking({
      intents,
      normalizations: [],
      from: Math.min(range.from, paragraph.start),
      to: Math.max(range.to, end),
    });
  }

  private resolve({
    operation,
    options,
    state,
    story,
    diff,
    tracked,
    suggested,
    mintList,
    besideDeletions,
  }: ResolveCanonicalPublicOperationOptions) {
    const refusal = (reason: FolioAIEditSkipReason, message?: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message: message ?? `The operation was refused: ${reason}.`,
        }),
      );
    const tableGap = (message: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason: "unsupportedBlock",
          message,
          gap: CANONICAL_GAP.publicTableStructure,
        }),
      );
    switch (operation.type) {
      case "replaceInBlock":
      case "replaceRange":
      case "replaceBlock":
      case "splitBlock":
      case "formatRange":
      case "mergeBlockWithNext":
      case "commentOnRange":
      case "commentOnBlock":
      case "deleteBlock":
      case "setBlockParagraphProperties":
      case "insertAfterBlock":
      case "insertBeforeBlock":
      case "insertTable":
        break;
      default:
        return refusal(
          "unsupportedBlock",
          `Canonical public operation ${operation.type} is unavailable.`,
        );
    }
    const blockId =
      operation.type === "replaceRange" ||
      operation.type === "formatRange" ||
      operation.type === "commentOnRange"
        ? operation.range.blockId
        : operation.blockId;
    const anchor = Object.hasOwn(options.snapshot.anchors, blockId)
      ? options.snapshot.anchors[blockId]
      : undefined;
    const projection = this.options.session.projectStory(story);
    if (!anchor || projection.isErr()) return refusal("missingBlock");
    const paragraph = projection.value.paragraph(blockId);
    if (!paragraph) return refusal("missingBlock");
    // A block beside a paragraph whose break is pending removal is placed
    // by the join it is part of (`besidePendingJoin`); any other operation
    // on it waits for that change to be resolved.
    if (
      (paragraph.source.pPrMark?.kind === "del" ||
        paragraph.source.pPrMark?.kind === "moveFrom") &&
      !placesBlock(operation)
    )
      return refusal("pendingDeletion");
    const noteReferences = collectNoteReferenceLabels(state.doc);
    const clean = buildCleanBlockText(paragraph.node, paragraph.start - 1, {
      fieldResults: "text",
      noteReferences,
    });
    const hash = hashFolioAIBlockContent(clean.text);
    if (operation.precondition && hash !== operation.precondition.blockTextHash)
      return refusal("preconditionFailed");
    if (hash !== anchor.textHash) return refusal("changedBlock");
    // A paragraph no paragraph follows, deleted directly by the same batch:
    // before it and after it are one place once it is gone, and after it,
    // the deletion keeps a paragraph to remove it by identity. A tracked
    // deletion leaves the paragraph in place, so its sides stay apart, and it
    // removes the paragraph by identity once the insertion beside it applies.
    const followed =
      state.doc.resolve(paragraph.start + paragraph.node.content.size + 1)
        .nodeAfter?.type.name === "paragraph";
    if (operation.type === "insertTable")
      return resolveTableInsertion({
        operation,
        paragraph,
        projection: projection.value,
        story,
        noteReferences,
        refusal,
        tableGap,
      });
    if (isBlockInsertion(operation)) {
      const besideDeletion =
        !tracked &&
        operation.type === "insertBeforeBlock" &&
        !followed &&
        besideDeletions.deleted.has(blockId);
      const insertion: BlockInsertion = besideDeletion
        ? { ...operation, type: "insertAfterBlock" }
        : operation;
      const placed = suggested
        ? besideAsItStands(paragraph)
        : besidePendingJoin({
            side: insertion.type === "insertAfterBlock" ? "after" : "before",
            paragraph,
            projection: projection.value,
            noteReferences,
          });
      if (placed === undefined) return refusal("pendingDeletion");
      const inserted = resolveBlockInsertion({
        operation: insertion,
        paragraph: placed.place,
        formattingSource: placed.formattingSource,
        projection: projection.value,
        story,
        mintList,
        refusal,
      });
      return besideDeletion && inserted.isOk()
        ? Result.ok({ ...inserted.value, besideDeletion: true as const })
        : inserted;
    }
    if (operation.type === "setBlockParagraphProperties")
      return resolveParagraphProperties({
        operation,
        paragraph,
        projection: projection.value,
        mintList,
        refusal,
      });
    if (operation.type === "deleteBlock") {
      const gainsFollower =
        !followed && besideDeletions.insertedBeside.has(blockId);
      const deleted = this.resolveBlockDeletion({
        operation,
        paragraph,
        projection: projection.value,
        projectionStory: story,
        doc: state.doc,
        tracked,
        gainsFollower,
        refusal,
      });
      return gainsFollower && deleted.isOk()
        ? Result.ok({ ...deleted.value, besideDeletion: true as const })
        : deleted;
    }
    let start = 0;
    let end = clean.text.length;
    let text = "";
    switch (operation.type) {
      case "replaceInBlock": {
        if (operation.find === "") return refusal("emptyOperation");
        start = clean.text.indexOf(operation.find);
        if (start < 0) return refusal("missingFind");
        if (clean.text.indexOf(operation.find, start + 1) >= 0)
          return refusal("ambiguousFind");
        end = start + operation.find.length;
        text = operation.replace;
        break;
      }
      case "commentOnBlock":
        if (operation.quote !== undefined) {
          start = clean.text.indexOf(operation.quote);
          if (start < 0) return refusal("missingFind");
          if (clean.text.indexOf(operation.quote, start + 1) >= 0)
            return refusal("ambiguousFind");
          end = start + operation.quote.length;
        }
        break;
      case "replaceRange":
      case "formatRange":
      case "commentOnRange":
        start = operation.range.startOffset;
        end = operation.range.endOffset;
        if (
          hashFolioAIBlockText(clean.text.slice(start, end)) !==
          operation.range.selectedTextHash
        )
          return refusal("staleRange");
        text = operation.type === "replaceRange" ? operation.replace : "";
        break;
      case "replaceBlock":
        if (
          operation.styleId !== undefined ||
          operation.preserveFormatting === false
        )
          return refusal(
            "unsupportedBlock",
            "Canonical replacement paragraph properties are unavailable.",
          );
        text = operation.text;
        break;
      case "splitBlock":
        start = operation.offset;
        end = start + (operation.separator?.length ?? 0);
        if (
          start <= 0 ||
          start >= clean.text.length ||
          clean.text.slice(start, end) !== (operation.separator ?? "")
        )
          return refusal("staleRange");
        break;
      case "mergeBlockWithNext":
        if (/[\t\r\n]/u.test(operation.separator ?? ""))
          return refusal(
            "unsupportedBlock",
            "Canonical public join separators require plain text.",
          );
        start = clean.text.length;
        end = start;
        break;
      default: {
        const unreachable: never = operation;
        return unreachable;
      }
    }
    if (
      [start, end].some(
        (offset) =>
          splitsSurrogatePair(clean.text, offset) ||
          splitsGraphemeCluster(clean.text, offset),
      )
    )
      return refusal("splitsCharacter");
    if (
      clean.structuralBoundaries.some(
        (boundary) =>
          boundary.type === "noteReference" &&
          start < boundary.offset + boundary.length &&
          end > boundary.offset,
      )
    )
      return refusal("protectedReference");
    if (stripInlineEmphasisMarkers(text) !== text)
      return refusal(
        "unsupportedBlock",
        "Canonical public emphasis replacements are unavailable.",
      );
    if (/\[\^(?:e)?\d+\]/u.test(text) || /[\t\r\n]/u.test(text))
      return refusal(
        "unsupportedBlock",
        "Canonical public replacements require plain text.",
      );
    const range = resolveCleanTextRange({
      cleanBlock: clean,
      startOffset: start,
      endOffset: end,
    });
    if (!range) return refusal("staleRange");
    const from = projection.value.addressAt(range.from);
    const to = projection.value.addressAt(range.to);
    if (from.isErr() || to.isErr()) return refusal("unsupportedBlock");
    const resolvedComment = resolveComment({
      operation,
      from: from.value,
      to: to.value,
    });
    if (
      operation.type === "commentOnBlock" ||
      operation.type === "commentOnRange"
    )
      return Result.ok({
        ...resolvedComment,
        intents: [],
        normalizations: [],
        from: range.from,
        to: range.to,
      });
    if (
      resolvedComment.type === "plain" &&
      operation.type !== "splitBlock" &&
      operation.type !== "formatRange" &&
      operation.type !== "mergeBlockWithNext" &&
      clean.text.slice(start, end) === text
    )
      return refusal("noopOperation");
    let intent: EditorIntent;
    switch (operation.type) {
      case "splitBlock":
        // The first half keeps the paragraph's identity, as the default
        // engine's split does.
        intent = {
          type: "splitParagraph",
          at: from.value,
          to: to.value,
          newHalf: "second",
          newBlockId: createEditorIntentIdAllocator({
            reservedBlockIds: this.options.session.reservedBlockIds,
          })(
            readModel(this.options.session.document),
            {
              type: "splitParagraph",
              at: from.value,
              to: to.value,
            },
            this.options.session.idLease,
          ).newBlockId,
        };
        break;
      case "mergeBlockWithNext": {
        const next = projection.value.nextParagraphInContainer(blockId);
        if (next === undefined) return refusal("missingBlock");
        // The merged paragraph keeps the first one's identity, as the
        // default engine's merge does.
        intent = {
          type: "joinParagraphs",
          story,
          blockId,
          nextBlockId: next.blockId,
          survivor: "first",
        };
        break;
      }
      case "formatRange": {
        const formatting = operation.formatting;
        const patch = {
          ...(formatting.bold === undefined ? {} : { bold: formatting.bold }),
          ...(formatting.italic === undefined
            ? {}
            : { italic: formatting.italic }),
          ...(formatting.underline === undefined
            ? {}
            : {
                underline:
                  formatting.underline === null
                    ? null
                    : { style: formatting.underline ? "single" : "none" },
              }),
          ...(formatting.strike === undefined
            ? {}
            : { strike: formatting.strike, doubleStrike: null }),
          ...(formatting.fontSizePt === undefined
            ? {}
            : {
                fontSize:
                  formatting.fontSizePt === null
                    ? null
                    : formatting.fontSizePt * 2,
              }),
          ...(formatting.fontFamily === undefined
            ? {}
            : {
                fontFamily:
                  formatting.fontFamily === null
                    ? null
                    : {
                        ascii: formatting.fontFamily,
                        hAnsi: formatting.fontFamily,
                      },
              }),
          ...(formatting.color === undefined
            ? {}
            : {
                color:
                  formatting.color === null
                    ? null
                    : {
                        rgb: formatting.color.replace(/^#/u, "").toUpperCase(),
                      },
              }),
        } as const satisfies RunPropsPatch;
        intent = { type: "formatRun", from: from.value, to: to.value, patch };
        break;
      }
      case "replaceInBlock":
      case "replaceRange":
      case "replaceBlock":
        intent = {
          type: "replaceText",
          from: from.value,
          to: to.value,
          text,
          ...(options.replacementBackground === "keep"
            ? {}
            : { runPropsPatch: { highlight: null, shading: null } }),
        };
        break;
      default: {
        const unreachable: never = operation;
        return panic(`Unknown canonical public operation ${unreachable}`);
      }
    }
    const intents: EditorIntent[] = [];
    const normalizations: FolioAIEditNormalization[] = [];
    if (intent.type === "replaceText") {
      let offset = start;
      let pending: { start: number; end: number; text: string } | null = null;
      const changes: { start: number; end: number; text: string }[] = [];
      for (const segment of diff(clean.text.slice(start, end), text)) {
        switch (segment.type) {
          case "equal":
            if (pending) changes.push(pending);
            pending = null;
            offset += segment.text.length;
            break;
          case "del":
            pending ??= { start: offset, end: offset, text: "" };
            offset += segment.text.length;
            pending.end = offset;
            break;
          case "ins":
            pending ??= { start: offset, end: offset, text: "" };
            pending.text += segment.text;
            break;
          default: {
            const unreachable: never = segment.type;
            return panic(`Unknown replacement diff ${unreachable}`);
          }
        }
      }
      if (pending) changes.push(pending);
      for (const change of changes.toReversed()) {
        const gap = resolveCleanTextRange({
          cleanBlock: clean,
          startOffset: change.start,
          endOffset: change.end,
        });
        if (!gap) return refusal("unsupportedBlock");
        if (change.text.length > 0 && gap.from < gap.to) {
          let first: readonly Mark[] | null = null;
          let uniform = true;
          paragraph.node.nodesBetween(
            gap.from - paragraph.start,
            gap.to - paragraph.start,
            (node) => {
              if (
                !node.isText ||
                node.marks.some((mark) => mark.type.name === "deletion")
              )
                return;
              const marks = node.marks.filter(
                (mark) => mark.type.name !== "insertion",
              );
              if (first === null) first = marks;
              else if (!Mark.sameSet(first, marks)) uniform = false;
            },
          );
          if (!uniform && normalizations.length === 0)
            normalizations.push({
              id: operation.id,
              code: "uniformReplacementFormatting",
            });
        }
        // A replaced range ends just after its last character, before what
        // the clean text skips there (a comment's reference mark, struck
        // text), since the new text follows the struck words.
        const lastCharacter =
          change.end > change.start ? clean.offsets[change.end - 1] : undefined;
        const changeEnd =
          lastCharacter === undefined
            ? gap.to
            : Math.max(gap.from, Math.min(gap.to, lastCharacter + 1));
        // A replaced range takes the editor's selection rule: its edges stay
        // inside the markers beside them, so replacing exactly the words a
        // comment covers keeps the comment on the new text, pending and
        // accepted. An insertion at a point takes the caret's gap.
        const edges =
          gap.from < changeEnd
            ? projection.value.rangeEdgesAt(gap.from, changeEnd)
            : undefined;
        const changeFrom = projection.value.addressAt(gap.from, edges?.start);
        const changeTo = projection.value.addressAt(changeEnd, edges?.end);
        if (changeFrom.isErr() || changeTo.isErr())
          return refusal("unsupportedBlock");
        intents.push({
          type: "replaceText",
          from: changeFrom.value,
          to: changeTo.value,
          text: change.text,
          ...(intent.runPropsPatch === undefined
            ? {}
            : { runPropsPatch: intent.runPropsPatch }),
        });
      }
    } else if (
      operation.type === "mergeBlockWithNext" &&
      intent.type === "joinParagraphs"
    ) {
      // The separator goes in at the first paragraph's end, before the join.
      const separator = operation.separator ?? "";
      if (separator !== "") {
        const at = projection.value.addressAt(
          paragraph.start + paragraph.node.content.size,
        );
        if (at.isErr()) return refusal("unsupportedBlock");
        intents.push({
          type: "replaceText",
          from: at.value,
          to: at.value,
          text: separator,
        });
      }
      intents.push(intent);
      // The merged paragraph's properties: on the first paragraph a direct
      // join keeps, on the second, whose mark a tracked join leaves.
      intents.push(
        ...statedPropertiesIntents({
          story,
          blockId: tracked ? intent.nextBlockId : intent.blockId,
          properties: operation.mergedParagraphProperties,
          mintList,
        }),
      );
    } else if (
      operation.type === "splitBlock" &&
      intent.type === "splitParagraph"
    ) {
      intents.push(intent);
      // Each half's properties, as the default engine sets them once the
      // split is made: the first half keeps the paragraph's identity, the
      // second takes the new one.
      for (const [half, properties] of [
        [paragraph.blockId, operation.firstParagraphProperties],
        [intent.newBlockId, operation.secondParagraphProperties],
      ] as const)
        intents.push(
          ...statedPropertiesIntents({
            story,
            blockId: half,
            properties,
            mintList,
          }),
        );
    } else intents.push(intent);
    const structural =
      operation.type === "splitBlock" ||
      operation.type === "mergeBlockWithNext";
    const nextParagraph =
      operation.type === "mergeBlockWithNext" &&
      intent.type === "joinParagraphs"
        ? projection.value.paragraph(intent.nextBlockId)
        : undefined;
    let claimedTo = range.to;
    if (structural) claimedTo = paragraph.start + paragraph.node.content.size;
    if (nextParagraph)
      claimedTo = nextParagraph.start + nextParagraph.node.content.size;
    return Result.ok({
      ...resolvedComment,
      intents,
      normalizations,
      from: structural ? paragraph.start : range.from,
      to: claimedTo,
    });
  }
}
