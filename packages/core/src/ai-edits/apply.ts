import type { Mark, Node as PMNode, Schema } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { expectRunPropertyChangeMarkAttrs } from "../prosemirror/attrs";
import { marksToTextFormatting } from "../prosemirror/conversion/fromProseDoc";
import { getFolioParaIdFromBlockId } from "../types/block-id";
import type { RunPropertyChange } from "../types/document";
import { stripBlockIdentityAttrs } from "./block-identity";
import { buildCleanBlockText } from "./clean-text";
import {
  hasInlineEmphasis,
  parseInlineEmphasisRuns,
  stripInlineEmphasisMarkers,
} from "./inline-emphasis";
import { hashFolioAIBlockText, normalizeFolioAIBlockText } from "./snapshot";
import {
  mergeTableRectangle,
  mergeTrackedVerticalTableCells,
  splitTableRectangle,
  splitTrackedVerticalTableCell,
} from "./table-cell-mutations";
import {
  planTableMutations,
  type TableMutationPlanTarget,
  type TableRectangle,
} from "./table-mutation-plan";
import {
  applyTableColumnDeletion,
  applyTableColumnInsertion,
  applyTableRowDeletion,
  applyTableRowInsertion,
  findTableColumnInsertion,
  findTableRowInsertion,
  getTableColumnCoordinateKey,
  type TableColumnDeletion,
  type TableColumnInsertion,
  type TableRowDeletion,
  type TableRowInsertion,
  type TableStructureRevision,
} from "./table-row-column-mutations";
import {
  findEnclosingTableBoundary,
  findEnclosingTableCell,
  findEnclosingTableRow,
  tableRectangleCutsMergedCell,
} from "./table-targets";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
  FolioAIEditSkipReason,
  FolioAIEditSkippedOperation,
  FolioAISignatureParty,
} from "./types";
import { diffWordSegments } from "./word-diff";

/**
 * The only editor surface the apply logic touches: a current `state`
 * and a `dispatch` that swaps in the next one. A live `EditorView`
 * satisfies this structurally, and so does a headless seam
 * (`{ state, dispatch: (tr) => { state = state.apply(tr); } }`) — the
 * apply path never reaches for anything DOM-bound on the view, which is
 * what lets the same operation applier drive both the React editor and
 * the server-side reviewer in `./headless`.
 */
export type FolioAIEditView = {
  state: EditorState;
  dispatch: (transaction: Transaction) => void;
};

type ApplyFolioAIEditOperationsOptions = {
  view: FolioAIEditView;
  snapshot: FolioAIEditSnapshot;
  operations: FolioAIEditOperation[];
  mode?: FolioAIEditApplyMode;
  author?: string;
  /** Optional author initials (w:initials) stamped alongside the author. */
  initials?: string;
  createCommentId?: (text: string) => number;
};

type ApplyFolioAIEditOperationsInternalOptions = ApplyFolioAIEditOperationsOptions & {
  revisionIdSeed?: number;
};

/**
 * Operation types applied in `"suggested"` mode. Every produced revision — an
 * inline mark, a whole-node `_suggestedInsert` marker, or a suggested
 * `trIns`/`trDel`/`cellMarker` — is stripped from serialized DOCX until
 * accepted. Cell merge/split (not row/column ops) and comment ops stay
 * `unsupportedMode`.
 */
const SUGGESTED_SUPPORTED_OPERATION_TYPES: ReadonlySet<FolioAIEditOperation["type"]> = new Set([
  "replaceInBlock",
  "replaceRange",
  "formatRange",
  "replaceBlock",
  "deleteBlock",
  "insertAfterBlock",
  "insertBeforeBlock",
  "insertSignatureTable",
  "insertTableRow",
  "deleteTableRow",
  "insertTableColumn",
  "deleteTableColumn",
]);

type ResolvedOperation = {
  operation: FolioAIEditOperation;
  from: number;
  to: number;
  blockFrom: number;
  blockTo: number;
  blockNode: PMNode;
  insertText?: string;
  tableRowInsertion?: TableRowInsertion;
  tableRowDeletion?: TableRowDeletion;
  tableColumnInsertion?: TableColumnInsertion;
  tableColumnDeletion?: TableColumnDeletion;
  tableCellMerge?: TableCellMerge;
  tableCellSplit?: TableCellSplit;
  commentId?: number;
  /**
   * Position in the input `operations` array, used as a secondary
   * sort key so same-position operations preserve the AI's logical
   * ordering when applied bottom-up.
   */
  originalIndex: number;
};

const applyReplaceBlockStyleId = ({
  item,
  tr,
}: {
  item: ResolvedOperation;
  tr: Transaction;
}): Transaction => {
  if (item.operation.type !== "replaceBlock" || item.operation.styleId === undefined) {
    return tr;
  }

  const block = tr.doc.nodeAt(item.blockFrom);
  if (!block) {
    return tr;
  }

  return tr.setNodeMarkup(item.blockFrom, undefined, {
    ...block.attrs,
    styleId: item.operation.styleId,
  });
};

type ApplyInlineFormattingOptions = {
  tr: Transaction;
  schema: Schema;
  from: number;
  to: number;
  formatting: Extract<FolioAIEditOperation, { type: "formatRange" }>["formatting"];
};

const applyInlineFormatting = ({
  tr,
  schema,
  from,
  to,
  formatting,
}: ApplyInlineFormattingOptions): Transaction => {
  for (const [name, enabled] of Object.entries(formatting)) {
    const markType = schema.marks[name];
    if (!markType) {
      continue;
    }
    if (enabled) {
      tr.addMark(
        from,
        to,
        name === "underline" ? markType.create({ style: "single" }) : markType.create(),
      );
      continue;
    }
    tr.removeMark(from, to, markType);
  }
  return tr;
};

const formattingWouldChange = (
  marks: readonly Mark[],
  formatting: Extract<FolioAIEditOperation, { type: "formatRange" }>["formatting"],
): boolean =>
  Object.entries(formatting).some(([name, enabled]) => {
    const mark = marks.find((candidate) => candidate.type.name === name);
    if (!enabled) {
      return mark !== undefined;
    }
    if (name === "underline") {
      return mark?.attrs["style"] !== "single";
    }
    return mark === undefined;
  });

type ApplyTrackedInlineFormattingOptions = ApplyInlineFormattingOptions & {
  doc: PMNode;
  revisionId: number;
  author: string;
  date: string;
  /** Non-null stamps the produced `runPropertyChange` mark as a suggestion. */
  suggestionId?: string | null;
};

const applyTrackedInlineFormatting = ({
  tr,
  schema,
  doc,
  from,
  to,
  formatting,
  revisionId,
  author,
  date,
  suggestionId = null,
}: ApplyTrackedInlineFormattingOptions): Transaction => {
  const propertyChangeType = schema.marks["runPropertyChange"];
  if (!propertyChangeType) {
    return tr;
  }

  const segments: { from: number; to: number; changes: RunPropertyChange[] }[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText || !formattingWouldChange(node.marks, formatting)) {
      return;
    }
    const segmentFrom = Math.max(from, pos);
    const segmentTo = Math.min(to, pos + node.nodeSize);
    const previousFormatting = marksToTextFormatting(node.marks);
    const existingMark = node.marks.find((mark) => mark.type === propertyChangeType);
    const existingChanges = existingMark
      ? expectRunPropertyChangeMarkAttrs(existingMark).changes
      : [];
    const change: RunPropertyChange = {
      type: "runPropertyChange",
      info: { id: revisionId, author, date },
      ...(Object.keys(previousFormatting).length > 0 ? { previousFormatting } : {}),
    };
    segments.push({
      from: segmentFrom,
      to: segmentTo,
      changes: [...existingChanges, change],
    });
  });

  if (segments.length === 0) {
    return tr;
  }
  const suggestionAttrs = suggestionId === null ? {} : { provenance: "suggested", suggestionId };
  applyInlineFormatting({ tr, schema, from, to, formatting });
  for (const segment of segments) {
    tr.addMark(
      segment.from,
      segment.to,
      propertyChangeType.create({ changes: segment.changes, ...suggestionAttrs }),
    );
  }
  return tr;
};

type LiveBlockEntry = { from: number; to: number; node: PMNode };

/**
 * Module-scoped monotonic counter for tracked-change revision ids.
 * Seeded once from `Date.now()` so ids are roughly time-ordered for
 * humans reading raw DOCX, then incremented per allocation. A bare
 * `Date.now()` seed per applyAIEditOperations call would collide
 * across batches that fire within the same millisecond (the panel's
 * Accept-all loop does exactly that — multiple calls in tight
 * succession). Reserving a contiguous range up front guarantees
 * uniqueness across overlapping calls in the same JS realm.
 */
let revisionIdCursor = Date.now() * 1000;
const nextRevisionSeed = (count: number): number => {
  // Each replace allocates two ids per op; insert/delete one each.
  // Reserve `count * 4` to be safely above any conceivable per-op
  // allocation (current max is 2). Returning the start of the
  // reserved range as the seed is enough — the caller bumps it.
  const start = revisionIdCursor;
  revisionIdCursor += Math.max(count, 1) * 4;
  return start;
};

/**
 * Walk the live doc once and bucket every textblock by its
 * normalised text hash. Resolution then maps each snapshot anchor
 * to the live block at the same ordinal among same-hash siblings —
 * unrelated edits that shift absolute positions no longer break
 * the lookup, and a sibling sharing text content with the target
 * doesn't trigger a false "changed" skip either.
 */
const collectLiveBlocksByHash = (doc: PMNode) => {
  const byHash = new Map<string, LiveBlockEntry[]>();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) {
      return;
    }
    // Hash from the post-tracked-changes view so the snapshot
    // (taken with the same view) and live doc bucket the same
    // block under the same key. Otherwise a block mid-edit gets a
    // different hash than the snapshot recorded and the resolver
    // skips it as "changed".
    const cleanText = buildCleanBlockText(node, pos).text;
    const hash = hashFolioAIBlockText(normalizeFolioAIBlockText(cleanText));
    const bucket = byHash.get(hash) ?? [];
    bucket.push({ from: pos, to: pos + node.nodeSize, node });
    byHash.set(hash, bucket);
  });
  return byHash;
};

/**
 * Index live textblocks by their `w14:paraId`. Used by the resolver
 * to prefer a paraId-anchored lookup when the snapshot id encodes
 * one. Direct lookup avoids
 * the hash+ordinal failure mode where an earlier-in-document
 * duplicate of the same text gets picked instead of the actual
 * referenced paragraph.
 */
const collectLiveBlocksByParaId = (doc: PMNode) => {
  const byParaId = new Map<string, LiveBlockEntry>();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) {
      return;
    }
    const paraId: unknown = node.attrs["paraId"];
    if (typeof paraId !== "string" || paraId.length === 0) {
      return;
    }
    // First-write-wins so an Enter-split duplicate (briefly co-existing
    // before the allocator re-issues) doesn't override the original.
    if (!byParaId.has(paraId)) {
      byParaId.set(paraId, { from: pos, to: pos + node.nodeSize, node });
    }
  });
  return byParaId;
};

type TableCellMerge = {
  tablePosition: number;
  rectangle: TableRectangle;
};

type TableCellSplit = TableCellMerge;

const getTableMutationPlanTarget = (item: ResolvedOperation): TableMutationPlanTarget => {
  if (item.tableCellMerge) {
    return { type: "mergeCells", ...item.tableCellMerge };
  }
  if (item.tableCellSplit) {
    return { type: "splitCell", ...item.tableCellSplit };
  }

  const tablePosition =
    item.tableRowInsertion?.tableStart !== undefined
      ? item.tableRowInsertion.tableStart - 1
      : (item.tableRowDeletion?.tablePosition ??
        item.tableColumnInsertion?.tablePosition ??
        item.tableColumnDeletion?.tablePosition);
  if (tablePosition !== undefined) {
    return { type: "tableStructure", tablePosition };
  }
  return { type: "none" };
};

/**
 * The snapshot recorded an `hashOccurrenceCount` per anchor but
 * not which ordinal within that bucket the block was — recompute
 * on demand from the snapshot's anchor map. Stable iteration
 * (object insertion order) means anchors with the same hash come
 * out in document order, which is what we want.
 */
const ordinalAmongSameHash = (snapshot: FolioAIEditSnapshot, blockId: string): number => {
  const target = snapshot.anchors[blockId];
  if (!target) {
    return -1;
  }
  let ordinal = 0;
  for (const anchor of Object.values(snapshot.anchors)) {
    if (anchor.id === blockId) {
      return ordinal;
    }
    if (anchor.textHash === target.textHash) {
      ordinal += 1;
    }
  }
  return -1;
};

/**
 * Length of the underscore signature rule. Mirrors the constant
 * in `docx-core/legal-source/compile.ts` so the on-screen signature
 * line matches what `create-document` produces.
 */
const SIGNATURE_LINE = "_".repeat(28);

type BuildSignatureTableNodeOptions = {
  schema: Schema;
  parties: readonly FolioAISignatureParty[];
};

/**
 * Build a borderless PM table mirroring docx-core's
 * `signatureTable` helper: one row, one cell per party, each cell
 * containing party name (bold), two spacer paragraphs, a signature
 * rule, then optional signatory and italic title lines.
 *
 * Returns `null` when the editor schema is missing one of the
 * required node types — callers should surface the op as a skip
 * rather than crashing.
 */
const buildSignatureTableNode = ({
  schema,
  parties,
}: BuildSignatureTableNodeOptions): PMNode | null => {
  const paragraphType = schema.nodes["paragraph"];
  const cellType = schema.nodes["tableCell"];
  const rowType = schema.nodes["tableRow"];
  const tableType = schema.nodes["table"];
  if (!paragraphType || !cellType || !rowType || !tableType) {
    return null;
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];

  const buildParagraph = (
    text: string,
    options: {
      styleId: string;
      bold?: boolean;
      italic?: boolean;
    },
  ): PMNode => {
    const marks: Mark[] = [];
    if (options.bold && boldType) {
      marks.push(boldType.create());
    }
    if (options.italic && italicType) {
      marks.push(italicType.create());
    }
    const content = text.length > 0 ? schema.text(text, marks) : null;
    return paragraphType.create({ styleId: options.styleId }, content);
  };

  const buildCell = (party: FolioAISignatureParty): PMNode => {
    const cellContent: PMNode[] = [
      buildParagraph(party.name, { styleId: "SignatureParty", bold: true }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph(SIGNATURE_LINE, { styleId: "SignatureRule" }),
    ];
    if (party.signatory && party.signatory.length > 0) {
      cellContent.push(buildParagraph(party.signatory, { styleId: "SignatureField" }));
    }
    if (party.title && party.title.length > 0) {
      cellContent.push(
        buildParagraph(party.title, {
          styleId: "SignatureField",
          italic: true,
        }),
      );
    }
    // Cell attrs left to schema defaults — column widths and
    // borders are decided by the table renderer; we just need the
    // structural shape.
    return cellType.create({}, cellContent);
  };

  const cells = parties.map(buildCell);
  const row = rowType.create({}, cells);
  return tableType.create({}, row);
};

/**
 * Build the inline content for an inserted or replaced block, promoting the
 * model's `**bold**` / `***bold italic***` markdown into real Word marks (the
 * edit-tool schema has no inline-format channel, so the model improvises with
 * markdown that would otherwise land as literal asterisks). Falls back to a
 * single verbatim text node when no emphasis is present, so plain prose is
 * never reshaped. `baseMarks` (insertion / comment) ride on every run.
 */
const buildEmphasisInlineContent = (
  schema: Schema,
  text: string,
  baseMarks: readonly Mark[],
): PMNode[] => {
  const runs = parseInlineEmphasisRuns(text);
  if (!runs.some((run) => run.bold || run.italic)) {
    return [schema.text(text, [...baseMarks])];
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];
  const nodes: PMNode[] = [];
  for (const run of runs) {
    if (run.text.length === 0) {
      continue;
    }
    const marks: Mark[] = [...baseMarks];
    if (run.bold && boldType) {
      marks.push(boldType.create());
    }
    if (run.italic && italicType) {
      marks.push(italicType.create());
    }
    nodes.push(schema.text(run.text, marks));
  }
  return nodes.length > 0 ? nodes : [schema.text(text, [...baseMarks])];
};

const applyFolioAIEditOperationsInternal = ({
  view,
  snapshot,
  operations,
  mode = "tracked-changes",
  author = "AI",
  initials,
  createCommentId,
  revisionIdSeed,
}: ApplyFolioAIEditOperationsInternalOptions): FolioAIEditApplyResult => {
  const applied: FolioAIEditAppliedOperation[] = [];
  const skipped: FolioAIEditSkippedOperation[] = [];
  const resolved: ResolvedOperation[] = [];
  const insertionType = view.state.schema.marks["insertion"];
  const deletionType = view.state.schema.marks["deletion"];
  const commentType = view.state.schema.marks["comment"];
  const claimedTableRows = new Set<string>();
  const claimedTableColumns = new Set<string>();

  // `"suggested"` is `"tracked-changes"` plus a provenance stamp, so every
  // tracked-change code path below keys off this rather than an exact
  // `=== "tracked-changes"` check.
  const producesTrackedChanges = mode !== "direct";
  const isSuggested = mode === "suggested";

  if (producesTrackedChanges && (!insertionType || !deletionType)) {
    return {
      applied,
      skipped: operations.map((operation) => ({
        id: operation.id,
        reason: "unsupportedBlock",
      })),
    };
  }

  // Build the live-block indexes once per batch so individual op
  // resolutions don't each re-walk the doc. ParaId-anchored ids are
  // resolved against `liveBlocksByParaId`; the hash bucket is only
  // the fallback for ordinal-encoded snapshot ids.
  const liveBlocks = collectLiveBlocksByHash(view.state.doc);
  const liveBlocksByParaId = collectLiveBlocksByParaId(view.state.doc);

  for (const [index, operation] of operations.entries()) {
    const commentText = getOperationCommentText(operation);
    if (commentText !== undefined && (!commentType || createCommentId === undefined)) {
      skipped.push({ id: operation.id, reason: "unsupportedBlock" });
      continue;
    }

    const resolution = resolveOperation({
      snapshot,
      operation,
      liveBlocks,
      liveBlocksByParaId,
      doc: view.state.doc,
    });
    if (resolution.type === "skip") {
      skipped.push({ id: operation.id, reason: resolution.reason });
      continue;
    }

    const deletion = resolution.operation.tableRowDeletion;
    if (deletion) {
      const rowKey = `${deletion.tablePosition}:${deletion.rowIndex}`;
      if (claimedTableRows.has(rowKey)) {
        skipped.push({ id: operation.id, reason: "noopOperation" });
        continue;
      }
      claimedTableRows.add(rowKey);
    }

    const columnDeletion = resolution.operation.tableColumnDeletion;
    if (columnDeletion) {
      const columnKey = getTableColumnCoordinateKey(columnDeletion);
      if (claimedTableColumns.has(columnKey)) {
        skipped.push({ id: operation.id, reason: "noopOperation" });
        continue;
      }
      claimedTableColumns.add(columnKey);
    }

    const commentId = commentText !== undefined ? createCommentId?.(commentText) : undefined;
    resolved.push({
      ...resolution.operation,
      originalIndex: index,
      ...(commentId !== undefined && { commentId }),
    });
  }

  const tablePlan = planTableMutations(
    resolved.map((item) => ({
      item,
      operationId: item.operation.id,
      target: getTableMutationPlanTarget(item),
    })),
  );
  skipped.push(...tablePlan.skipped);
  const executableResolved = tablePlan.executable;

  if (executableResolved.length === 0) {
    return { applied, skipped };
  }

  let tr = view.state.tr;
  let revisionSeed = revisionIdSeed ?? nextRevisionSeed(executableResolved.length);
  const date = new Date().toISOString();
  const insertedColumnCounts = new Map<string, number>();

  // Sort right-to-left so each tr.insert / tr.delete leaves earlier
  // positions intact. Column insertions run before deletions at the
  // same snapshot coordinate; the deletion path accounts for those
  // inserted columns so it still removes the original target. Other
  // ties use reverse input order so repeated insertions retain their
  // requested sequence.
  for (const item of executableResolved.toSorted((left, right) => {
    const leftCellShape = left.tableCellMerge ?? left.tableCellSplit;
    const rightCellShape = right.tableCellMerge ?? right.tableCellSplit;
    if (!leftCellShape && rightCellShape) {
      return -1;
    }
    if (leftCellShape && !rightCellShape) {
      return 1;
    }
    if (leftCellShape && rightCellShape) {
      if (leftCellShape.tablePosition !== rightCellShape.tablePosition) {
        return rightCellShape.tablePosition - leftCellShape.tablePosition;
      }
      if (leftCellShape.rectangle.bottom !== rightCellShape.rectangle.bottom) {
        return rightCellShape.rectangle.bottom - leftCellShape.rectangle.bottom;
      }
      if (leftCellShape.rectangle.right !== rightCellShape.rectangle.right) {
        return rightCellShape.rectangle.right - leftCellShape.rectangle.right;
      }
      return right.originalIndex - left.originalIndex;
    }
    const leftColumn = left.tableColumnInsertion ?? left.tableColumnDeletion;
    const rightColumn = right.tableColumnInsertion ?? right.tableColumnDeletion;
    if (!leftColumn && rightColumn) {
      return -1;
    }
    if (leftColumn && !rightColumn) {
      return 1;
    }
    if (leftColumn && rightColumn) {
      if (leftColumn.tablePosition !== rightColumn.tablePosition) {
        return rightColumn.tablePosition - leftColumn.tablePosition;
      }
      if (leftColumn.columnIndex !== rightColumn.columnIndex) {
        return rightColumn.columnIndex - leftColumn.columnIndex;
      }
      const leftIsInsertion = left.tableColumnInsertion !== undefined;
      const rightIsInsertion = right.tableColumnInsertion !== undefined;
      if (leftIsInsertion !== rightIsInsertion) {
        return leftIsInsertion ? -1 : 1;
      }
      return right.originalIndex - left.originalIndex;
    }
    if (left.from !== right.from) {
      return right.from - left.from;
    }
    return right.originalIndex - left.originalIndex;
  })) {
    const commentMark =
      item.commentId !== undefined && commentType
        ? commentType.create({ commentId: item.commentId })
        : null;

    // Snapshot the transaction's step count so we can detect when an
    // operation produced zero document changes and report it as a
    // skipped no-op instead of a phantom "applied" entry. This caught
    // the silent accept-failure bug where a replaceInBlock on a
    // block with pending tracked changes computed the wrong source
    // text and the diff produced no marks; the panel said "accepted"
    // but the doc was untouched.
    const stepsBefore = tr.steps.length;
    let appliedRevisionIds: number[] | undefined;

    // Suggested mode covers the inline text/format operations plus the block and
    // table row/column structural operations (see
    // `SUGGESTED_SUPPORTED_OPERATION_TYPES`). Every revision they produce — an
    // inline mark, a whole-node `_suggestedInsert` marker, or a suggested
    // `trIns`/`trDel`/`cellMarker` — is removed by the serialization strip.
    // Operations outside the allowlist (comment ops, cell merge/split) report
    // `unsupportedMode` so no unstrippable suggested state can be produced.
    if (isSuggested && !SUGGESTED_SUPPORTED_OPERATION_TYPES.has(item.operation.type)) {
      skipped.push({ id: item.operation.id, reason: "unsupportedMode" });
      continue;
    }

    // Every mark a suggested operation produces is stamped with this id so the
    // host can accept/reject the whole suggestion at once. Falls back to the
    // operation id when the caller does not supply one.
    const suggestionId: string | null = isSuggested
      ? (item.operation.suggestionId ?? item.operation.id)
      : null;

    // Fields merged into every node-attr revision (trIns/trDel/cellMarker) and
    // whole-node `_suggestedInsert` marker this operation produces.
    const trackedRevisionExtras = {
      ...(initials ? { initials } : {}),
      ...(suggestionId !== null ? { provenance: "suggested" as const, suggestionId } : {}),
    };
    // The structural revision a table op writes. Uses the current `revisionSeed`
    // (each table branch consumes it with `revisionSeed++` after applying); one
    // op runs per iteration, so no prior increment shifts this value.
    const structuralRevision: TableStructureRevision | null = producesTrackedChanges
      ? { revisionId: revisionSeed, author, date, ...trackedRevisionExtras }
      : null;

    switch (item.operation.type) {
      case "replaceInBlock":
      case "replaceRange": {
        const revisionIdDelete = revisionSeed++;
        const revisionIdInsert = revisionSeed++;
        tr = applyTextReplacement({
          tr,
          item,
          mode,
          author,
          date,
          revisionIdDelete,
          revisionIdInsert,
          commentMark,
          suggestionId,
          initials,
        });
        if (producesTrackedChanges) {
          appliedRevisionIds = [revisionIdDelete, revisionIdInsert];
        }
        break;
      }
      case "commentOnRange": {
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      case "formatRange": {
        if (producesTrackedChanges) {
          const revisionId = revisionSeed++;
          tr = applyTrackedInlineFormatting({
            tr,
            schema: view.state.schema,
            doc: tr.doc,
            from: item.from,
            to: item.to,
            formatting: item.operation.formatting,
            revisionId,
            author,
            date,
            suggestionId,
          });
          appliedRevisionIds = [revisionId];
          break;
        }
        tr = applyInlineFormatting({
          tr,
          schema: view.state.schema,
          from: item.from,
          to: item.to,
          formatting: item.operation.formatting,
        });
        break;
      }
      case "replaceBlock": {
        const revisionIdDelete = revisionSeed++;
        const revisionIdInsert = revisionSeed++;
        // Default to preserving formatting (existing behaviour);
        // when explicitly disabled and we're in direct mode, swap
        // the whole block node for a fresh paragraph that drops
        // all block-level attrs. tracked-changes mode keeps the
        // attrs because the visible diff is text-only.
        if (item.operation.preserveFormatting === false && mode === "direct") {
          const replacement = item.operation.text;
          const paragraphType = view.state.schema.nodes["paragraph"];
          if (paragraphType) {
            const replaceAttrs: Record<string, unknown> | null =
              item.operation.styleId !== undefined ? { styleId: item.operation.styleId } : null;
            const node = paragraphType.create(
              replaceAttrs,
              replacement.length === 0
                ? null
                : buildEmphasisInlineContent(view.state.schema, replacement, []),
            );
            tr = tr.replaceWith(item.blockFrom, item.blockTo, node);
            break;
          }
        }
        // Direct mode, formatting preserved (the default): when the replacement
        // carries inline emphasis, rebuild the block node keeping its own attrs
        // so `**bold**` becomes real marks. The tracked-changes path can't carry
        // inline marks through its word-diff redline, so it still strips them;
        // plain replacements fall through to the text-only swap below unchanged.
        if (mode === "direct" && hasInlineEmphasis(item.operation.text)) {
          const attrs: Record<string, unknown> =
            item.operation.styleId !== undefined
              ? { ...item.blockNode.attrs, styleId: item.operation.styleId }
              : { ...item.blockNode.attrs };
          const node = item.blockNode.type.create(
            attrs,
            buildEmphasisInlineContent(
              view.state.schema,
              item.operation.text,
              commentMark ? [commentMark] : [],
            ),
          );
          tr = tr.replaceWith(item.blockFrom, item.blockTo, node);
          break;
        }
        tr = applyTextReplacement({
          tr,
          item,
          mode,
          author,
          date,
          revisionIdDelete,
          revisionIdInsert,
          commentMark,
          suggestionId,
          initials,
        });
        tr = applyReplaceBlockStyleId({ item, tr });
        if (producesTrackedChanges) {
          appliedRevisionIds = [revisionIdDelete, revisionIdInsert];
        }
        break;
      }
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        if (
          mode === "tracked-changes" &&
          item.operation.pageBreakBefore === true &&
          (!item.insertText || item.insertText.length === 0)
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedMode",
          });
          continue;
        }

        const marks = [];
        let insertedBlockRevisionId: number | null = null;
        if (producesTrackedChanges && insertionType) {
          const revisionId = revisionSeed++;
          insertedBlockRevisionId = revisionId;
          marks.push(
            insertionType.create({
              revisionId,
              author,
              date,
              ...trackedRevisionExtras,
            }),
          );
          appliedRevisionIds = [revisionId];
        }
        if (commentMark) {
          marks.push(commentMark);
        }
        const content =
          item.insertText && item.insertText.length > 0
            ? buildEmphasisInlineContent(view.state.schema, item.insertText, marks)
            : null;
        // Inherit formatting attrs (listMarker, styleId, …) from
        // the source block but never reuse identity attrs — a new
        // paragraph must get fresh paraId/textId so trackers don't
        // collide.
        const baseAttrs =
          item.operation.inheritFormatting === false
            ? {}
            : stripBlockIdentityAttrs(item.blockNode.attrs);
        const attrs: Record<string, unknown> = { ...baseAttrs };
        if (item.operation.pageBreakBefore === true) {
          attrs["pageBreakBefore"] = true;
        }
        if (item.operation.styleId !== undefined) {
          // Heading / clause style ids (e.g. ClauseHeading1) take
          // precedence over the source block's style so the
          // inserted paragraph renders as the requested kind, not
          // as a clone of the anchor.
          attrs["styleId"] = item.operation.styleId;
          // A heading is logically a fresh block; drop list marker
          // attrs that would otherwise leak from the anchor and
          // render the heading as a list item.
          if (item.operation.inheritFormatting !== false) {
            attrs["listMarker"] = null;
            attrs["listMarkerHidden"] = null;
            attrs["listLevelNumFmts"] = null;
            attrs["listLevelStarts"] = null;
            attrs["listAbstractNumId"] = null;
            attrs["listStartOverride"] = null;
          }
        }
        // In suggested mode, mark the whole inserted paragraph so the strip
        // drops it from serialized DOCX until accepted (the inline insertion
        // marks alone would leave an empty paragraph behind).
        if (isSuggested && suggestionId !== null && insertedBlockRevisionId !== null) {
          attrs["_suggestedInsert"] = {
            suggestionId,
            revisionId: insertedBlockRevisionId,
            author,
            date,
            ...(initials ? { initials } : {}),
          };
        }
        const node = item.blockNode.type.create(attrs, content);
        tr = tr.insert(item.from, node);
        break;
      }
      case "insertSignatureTable": {
        // Tracked-changes mode has no whole-table-insert primitive in OOXML, so
        // it stays unsupported. Suggested mode CAN carry it: the whole table is
        // flagged `_suggestedInsert` and stripped until accepted (accepting
        // applies it directly — see acceptSuggestion).
        if (mode === "tracked-changes") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedMode",
          });
          continue;
        }

        const signatureTable = buildSignatureTableNode({
          schema: view.state.schema,
          parties: item.operation.parties,
        });
        if (!signatureTable) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        // Direct mode: tables don't carry tracked-change marks — the table
        // structure itself is the insert, and inline insertion marks on the
        // paragraph runs inside would double-up with the structural addition.
        let node = signatureTable;
        if (isSuggested && suggestionId !== null) {
          const revisionId = revisionSeed++;
          node = signatureTable.type.create(
            {
              ...signatureTable.attrs,
              _suggestedInsert: {
                suggestionId,
                revisionId,
                author,
                date,
                ...(initials ? { initials } : {}),
              },
            },
            signatureTable.content,
            signatureTable.marks,
          );
          appliedRevisionIds = [revisionId];
        }
        tr = tr.insert(item.from, node);
        break;
      }
      case "insertTableRow": {
        const insertion = item.tableRowInsertion;
        if (!insertion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableRowInsertion({
          tr,
          insertion,
          cellTexts: item.operation.cellTexts,
          revision,
        });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        break;
      }
      case "insertTableColumn": {
        const insertion = item.tableColumnInsertion;
        if (!insertion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableColumnInsertion({
          tr,
          insertion,
          cellTexts: item.operation.cellTexts,
          revision,
        });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        const columnKey = getTableColumnCoordinateKey(insertion);
        insertedColumnCounts.set(columnKey, (insertedColumnCounts.get(columnKey) ?? 0) + 1);
        break;
      }
      case "deleteTableColumn": {
        const deletion = item.tableColumnDeletion;
        if (!deletion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const columnKey = getTableColumnCoordinateKey(deletion);
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableColumnDeletion({
          tr,
          deletion,
          insertedColumnCount: insertedColumnCounts.get(columnKey) ?? 0,
          revision,
        });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        break;
      }
      case "mergeTableCells": {
        const merge = item.tableCellMerge;
        const tablePosition = merge ? tr.mapping.map(merge.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !merge ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed : null;
        const nextTr =
          revisionId === null
            ? mergeTableRectangle({ tr, tablePosition, table, rectangle: merge.rectangle })
            : mergeTrackedVerticalTableCells({
                tr,
                tablePosition,
                table,
                rectangle: merge.rectangle,
                revisionId,
                author,
                date,
              });
        if (!nextTr) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = nextTr;
        if (revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [revisionId];
        }
        break;
      }
      case "splitTableCell": {
        const split = item.tableCellSplit;
        const tablePosition = split ? tr.mapping.map(split.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !split ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed : null;
        const nextTr =
          revisionId === null
            ? splitTableRectangle({ tr, tablePosition, table, rectangle: split.rectangle })
            : splitTrackedVerticalTableCell({
                tr,
                tablePosition,
                table,
                rectangle: split.rectangle,
                revisionId,
                author,
                date,
              });
        if (!nextTr) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = nextTr;
        if (revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [revisionId];
        }
        break;
      }
      case "deleteTableRow": {
        const deletion = item.tableRowDeletion;
        if (!deletion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableRowDeletion({ tr, deletion, revision });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        break;
      }
      case "deleteBlock": {
        if (mode === "direct") {
          tr = tr.delete(item.blockFrom, item.blockTo);
          break;
        }

        if (deletionType) {
          const revisionId = revisionSeed++;
          tr = tr.addMark(
            item.from,
            item.to,
            deletionType.create({
              revisionId,
              author,
              date,
              ...trackedRevisionExtras,
            }),
          );
          appliedRevisionIds = [revisionId];
        }
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      case "commentOnBlock": {
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      default:
        break;
    }

    // commentOnBlock is intentionally a doc-mutation-free op
    // (adds a mark; that DOES count as a step) but covers the
    // edge case where the comment mark is missing. Treat any op
    // that emitted zero transaction steps as a no-op skip.
    if (tr.steps.length === stepsBefore) {
      skipped.push({ id: item.operation.id, reason: "noopOperation" });
      continue;
    }

    // Surface the primary id (first one) on the legacy `revisionId`
    // field so callers that just need a stable scroll/visual
    // reference keep working. The full set is on `revisionIds` for
    // accept/reject paths that must clear every mark belonging to
    // this op.
    applied.push({
      id: item.operation.id,
      ...(item.commentId !== undefined && { commentId: item.commentId }),
      ...(appliedRevisionIds !== undefined &&
        appliedRevisionIds[0] !== undefined && {
          revisionId: appliedRevisionIds[0],
          revisionIds: appliedRevisionIds,
        }),
      ...(suggestionId !== null && { suggestionId }),
    });
  }

  if (tr.docChanged) {
    view.dispatch(tr);
  }

  return { applied, skipped };
};

export const applyFolioAIEditOperations = (
  options: ApplyFolioAIEditOperationsOptions,
): FolioAIEditApplyResult => applyFolioAIEditOperationsInternal(options);

export const previewFolioAIEditOperations = (
  options: ApplyFolioAIEditOperationsOptions,
): FolioAIEditApplyResult => {
  const { view, createCommentId, ...applyOptions } = options;
  let previewCommentId = -1;
  const previewView: FolioAIEditView = {
    state: view.state,
    dispatch: (transaction) => {
      previewView.state = previewView.state.apply(transaction);
    },
  };
  const result = applyFolioAIEditOperationsInternal({
    ...applyOptions,
    view: previewView,
    ...(createCommentId !== undefined && {
      createCommentId: () => previewCommentId--,
    }),
    revisionIdSeed: -1_000_000_000,
  });
  return {
    applied: result.applied.map(({ id }) => ({ id })),
    skipped: result.skipped,
  };
};

type TextReplacementOptions = {
  tr: Transaction;
  item: ResolvedOperation;
  mode: FolioAIEditApplyMode;
  author: string;
  date: string;
  /**
   * Distinct revision id used for the deletion-side marks.
   * fromProseDoc treats a single revisionId carrying BOTH ins and
   * del marks as a Word "moveTo/moveFrom" pair on serialization,
   * which is wrong for an AI replace — so the engine allocates one
   * id for the deletion side and a separate one for the insertion
   * side of the same operation.
   */
  revisionIdDelete: number;
  /** Distinct revision id used for the insertion-side marks. */
  revisionIdInsert: number;
  commentMark: Mark | null;
  /** Non-null stamps every produced insertion/deletion mark as a suggestion. */
  suggestionId?: string | null;
  /** Optional author initials stamped on the produced marks. */
  initials?: string | undefined;
};

const applyTextReplacement = ({
  tr,
  item,
  mode,
  author,
  date,
  revisionIdDelete,
  revisionIdInsert,
  commentMark,
  suggestionId = null,
  initials,
}: TextReplacementOptions): Transaction => {
  let nextTr = tr;
  const replacement = stripInlineEmphasisMarkers(
    (() => {
      if (item.operation.type === "replaceInBlock" || item.operation.type === "replaceRange") {
        return item.operation.replace;
      }
      if (item.operation.type === "replaceBlock") {
        return item.operation.text;
      }
      return "";
    })(),
  );

  if (mode === "direct") {
    // Partial-block replacement carrying inline emphasis: rebuild the matched
    // range as real bold/italic runs instead of stripping the markers to plain
    // text. Full-block replaceBlock in direct mode is intercepted earlier and
    // never reaches here with emphasis; the tracked-changes path below still
    // strips, since its word-diff redline can't carry inline marks.
    if (
      (item.operation.type === "replaceInBlock" || item.operation.type === "replaceRange") &&
      hasInlineEmphasis(item.operation.replace)
    ) {
      const content = buildEmphasisInlineContent(
        nextTr.doc.type.schema,
        item.operation.replace,
        commentMark ? [commentMark] : [],
      );
      return nextTr.replaceWith(item.from, item.to, content);
    }
    nextTr = nextTr.insertText(replacement, item.from, item.to);
    if (commentMark && replacement.length > 0) {
      nextTr = nextTr.addMark(item.from, item.from + replacement.length, commentMark);
    }
    return nextTr;
  }

  const insertionType = nextTr.doc.type.schema.marks["insertion"];
  const deletionType = nextTr.doc.type.schema.marks["deletion"];
  const suggestionAttrs = suggestionId === null ? {} : { provenance: "suggested", suggestionId };
  const initialsAttr = initials ? { initials } : {};
  const delAttrs = {
    revisionId: revisionIdDelete,
    author,
    date,
    ...initialsAttr,
    ...suggestionAttrs,
  };
  const insAttrs = {
    revisionId: revisionIdInsert,
    author,
    date,
    ...initialsAttr,
    ...suggestionAttrs,
  };

  // Word-level diff is only safe when the source range maps to PM
  // positions losslessly. The block must have no atomic inline
  // nodes (hard breaks, inline images) — those break textContent /
  // PM-position alignment in ways the offsets array can't resolve.
  //
  // Existing tracked-change marks ARE handled here: we walk PM
  // positions through `cleanBlock.offsets[]` (built from the
  // post-tracked-changes view), so each clean-text char anchors at
  // the right live position even when the block has pending
  // deletion runs interleaved between surviving chars. Naively
  // accumulating `cursor += seg.text.length` would skip the gap
  // introduced by those deletion runs and write marks onto the
  // wrong live characters — the silent accept-failure bug.
  const blockHasOnlyTextChildren =
    item.blockNode.content.size === item.blockNode.textContent.length;
  const cleanBlock = blockHasOnlyTextChildren
    ? buildCleanBlockText(item.blockNode, item.blockFrom)
    : null;
  let sourceText: string | null = null;
  let sourceCleanStart = 0;
  if (cleanBlock !== null) {
    if (item.operation.type === "replaceInBlock") {
      sourceText = item.operation.find;
      sourceCleanStart = cleanBlock.text.indexOf(item.operation.find);
      if (sourceCleanStart === -1) {
        sourceText = null;
      }
    } else if (item.operation.type === "replaceRange") {
      sourceCleanStart = item.operation.range.startOffset;
      sourceText = cleanBlock.text.slice(sourceCleanStart, item.operation.range.endOffset);
    } else if (item.operation.type === "replaceBlock") {
      sourceText = cleanBlock.text;
      sourceCleanStart = 0;
    }
  }

  if (sourceText !== null && cleanBlock !== null) {
    const segments = diffWordSegments(sourceText, replacement);
    const offsets = cleanBlock.offsets;
    const offsetAt = (cleanOffset: number): number | null => offsets[cleanOffset] ?? null;

    type Step =
      | { kind: "del"; from: number; to: number }
      | { kind: "ins"; at: number; text: string };
    const steps: Step[] = [];
    // Cursor walks SOURCE-text offsets (within `sourceText`), then
    // we translate to PM positions through offsets[]. This survives
    // gaps caused by existing deletion-marked runs in the live doc.
    let cursor = 0;
    let allPositionsResolved = true;
    for (const seg of segments) {
      if (seg.type === "equal") {
        cursor += seg.text.length;
        continue;
      }
      if (seg.type === "del") {
        const pmFrom = offsetAt(sourceCleanStart + cursor);
        const pmTo = offsetAt(sourceCleanStart + cursor + seg.text.length);
        if (pmFrom === null || pmTo === null) {
          allPositionsResolved = false;
          break;
        }
        steps.push({ kind: "del", from: pmFrom, to: pmTo });
        cursor += seg.text.length;
        continue;
      }
      const pmAt = offsetAt(sourceCleanStart + cursor);
      if (pmAt === null) {
        allPositionsResolved = false;
        break;
      }
      steps.push({ kind: "ins", at: pmAt, text: seg.text });
    }

    if (allPositionsResolved) {
      // Apply right-to-left so earlier steps' source positions stay
      // valid after later steps mutate the doc.
      for (const step of steps.toReversed()) {
        if (step.kind === "del" && deletionType) {
          nextTr = nextTr.addMark(step.from, step.to, deletionType.create(delAttrs));
          if (commentMark) {
            nextTr = nextTr.addMark(step.from, step.to, commentMark);
          }
          continue;
        }
        if (step.kind === "ins" && insertionType) {
          nextTr = nextTr.insertText(step.text, step.at, step.at);
          nextTr = nextTr.addMark(
            step.at,
            step.at + step.text.length,
            insertionType.create(insAttrs),
          );
          if (commentMark) {
            nextTr = nextTr.addMark(step.at, step.at + step.text.length, commentMark);
          }
        }
      }
      return nextTr;
    }
    // else fall through to the coarse del+ins path below: the
    // offsets array didn't cover one of our boundaries, which only
    // happens for edge cases at the trailing block boundary.
  }

  if (replacement.length > 0 && insertionType) {
    nextTr = nextTr.insertText(replacement, item.to, item.to);
    nextTr = nextTr.addMark(item.to, item.to + replacement.length, insertionType.create(insAttrs));
    if (commentMark) {
      nextTr = nextTr.addMark(item.to, item.to + replacement.length, commentMark);
    }
  }

  if (item.to > item.from && deletionType) {
    nextTr = nextTr.addMark(item.from, item.to, deletionType.create(delAttrs));
    if (commentMark && replacement.length === 0) {
      nextTr = nextTr.addMark(item.from, item.to, commentMark);
    }
  }

  return nextTr;
};

type ResolvedBase = Omit<ResolvedOperation, "originalIndex" | "commentId">;

type ResolveOperationArgs = {
  snapshot: FolioAIEditSnapshot;
  operation: FolioAIEditOperation;
  liveBlocks: Map<string, LiveBlockEntry[]>;
  liveBlocksByParaId: Map<string, LiveBlockEntry>;
  doc: PMNode;
};

type ResolveStableBlockArgs = Omit<ResolveOperationArgs, "operation" | "doc"> & {
  blockId: string;
};

type StableBlockResolution = {
  type: "resolved";
  blockNode: PMNode;
  blockFrom: number;
  blockTo: number;
  cleanBlock: ReturnType<typeof buildCleanBlockText>;
  currentText: string;
  currentTextHash: string;
};

const resolveStableBlock = ({
  snapshot,
  blockId,
  liveBlocks,
  liveBlocksByParaId,
}: ResolveStableBlockArgs): StableBlockResolution | OperationResolutionSkip => {
  const anchor = snapshot.anchors[blockId];
  if (!anchor) {
    return { type: "skip", reason: "missingBlock" };
  }

  const encodedParaId = getFolioParaIdFromBlockId(blockId);
  let live: LiveBlockEntry | undefined;
  if (encodedParaId !== null) {
    live = liveBlocksByParaId.get(encodedParaId);
    if (!live) {
      return { type: "skip", reason: "missingBlock" };
    }
  } else {
    const ordinal = ordinalAmongSameHash(snapshot, blockId);
    if (ordinal < 0) {
      return { type: "skip", reason: "missingBlock" };
    }
    live = liveBlocks.get(anchor.textHash)?.[ordinal];
  }
  if (!live || !live.node.isTextblock) {
    return { type: "skip", reason: "changedBlock" };
  }

  const cleanBlock = buildCleanBlockText(live.node, live.from);
  const currentText = cleanBlock.text;
  const currentTextHash = hashFolioAIBlockText(normalizeFolioAIBlockText(currentText));
  if (currentTextHash !== anchor.textHash) {
    return { type: "skip", reason: "changedBlock" };
  }
  return {
    type: "resolved",
    blockNode: live.node,
    blockFrom: live.from,
    blockTo: live.to,
    cleanBlock,
    currentText,
    currentTextHash,
  };
};

const resolveOperation = ({
  snapshot,
  operation,
  liveBlocks,
  liveBlocksByParaId,
  doc,
}: ResolveOperationArgs):
  | { type: "resolved"; operation: ResolvedBase }
  | OperationResolutionSkip => {
  const blockId =
    operation.type === "replaceRange" ||
    operation.type === "commentOnRange" ||
    operation.type === "formatRange"
      ? operation.range.blockId
      : operation.blockId;
  const primaryBlock = resolveStableBlock({
    snapshot,
    blockId,
    liveBlocks,
    liveBlocksByParaId,
  });
  if (primaryBlock.type === "skip") {
    return primaryBlock;
  }
  const { blockNode, blockFrom, blockTo, cleanBlock, currentText, currentTextHash } = primaryBlock;
  if (
    operation.precondition !== undefined &&
    currentTextHash !== operation.precondition.blockTextHash
  ) {
    return { type: "skip", reason: "preconditionFailed" };
  }

  if (
    operation.type === "replaceRange" ||
    operation.type === "commentOnRange" ||
    operation.type === "formatRange"
  ) {
    const { startOffset, endOffset, selectedTextHash } = operation.range;
    const from = cleanBlock.offsets[startOffset];
    const to = cleanBlock.offsets[endOffset];
    if (from === undefined || to === undefined) {
      return { type: "skip", reason: "staleRange" };
    }
    const selectedText = currentText.slice(startOffset, endOffset);
    if (hashFolioAIBlockText(selectedText) !== selectedTextHash) {
      return { type: "skip", reason: "staleRange" };
    }
    if (operation.type === "replaceRange" && selectedText === operation.replace) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from,
        to,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "insertAfterBlock" || operation.type === "insertBeforeBlock") {
    // An empty `text` is normally an error, but a page-break-only
    // paragraph is legitimate (the user just wants whitespace +
    // forced break). Letting it through with no inline content
    // means the inserted block renders as an empty paragraph that
    // starts a new page.
    if (operation.text.length === 0 && operation.pageBreakBefore !== true) {
      return { type: "skip", reason: "emptyOperation" };
    }
    // If the anchor lives inside a `tableCell`, the model meant
    // "place the new block adjacent to the table", not "stuff it
    // into the cell". Override the insertion bounds to the table's
    // outer boundary so the synthesized sibling lands as a peer of
    // the table at doc level.
    const tableBoundary = findEnclosingTableBoundary(doc, blockFrom);
    const isInsertAfter = operation.type === "insertAfterBlock";
    let insertFrom: number;
    if (tableBoundary) {
      insertFrom = isInsertAfter ? tableBoundary.after : tableBoundary.before;
    } else {
      insertFrom = isInsertAfter ? blockTo : blockFrom;
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertFrom,
        to: insertFrom,
        blockFrom,
        blockTo,
        blockNode,
        insertText: operation.text,
      },
    };
  }

  if (operation.type === "insertSignatureTable") {
    if (operation.parties.length === 0) {
      return { type: "skip", reason: "emptyOperation" };
    }
    const position = operation.position ?? "after";
    const tableBoundary = findEnclosingTableBoundary(doc, blockFrom);
    let insertFrom: number;
    if (tableBoundary) {
      insertFrom = position === "after" ? tableBoundary.after : tableBoundary.before;
    } else {
      insertFrom = position === "after" ? blockTo : blockFrom;
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertFrom,
        to: insertFrom,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "insertTableRow") {
    const position = operation.position ?? "after";
    const insertion = findTableRowInsertion({ doc, blockFrom, position });
    if (!insertion || (operation.cellTexts?.length ?? 0) > insertion.cells.length) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertion.rowPosition,
        to: insertion.rowPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableRowInsertion: insertion,
      },
    };
  }

  if (operation.type === "deleteTableRow") {
    const target = findEnclosingTableRow(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.rowPosition,
        to: target.rowPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableRowDeletion: {
          tableStart: target.tableStart,
          tablePosition: target.tablePosition,
          rowIndex: target.rowIndex,
          rowPosition: target.rowPosition,
        },
      },
    };
  }

  if (operation.type === "insertTableColumn") {
    const position = operation.position ?? "after";
    const insertion = findTableColumnInsertion({ doc, blockFrom, position });
    if (!insertion) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const { boundaryPosition, ...tableColumnInsertion } = insertion;
    return {
      type: "resolved",
      operation: {
        operation,
        from: boundaryPosition,
        to: boundaryPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableColumnInsertion,
      },
    };
  }

  if (operation.type === "deleteTableColumn") {
    const target = findEnclosingTableCell(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.cellPosition,
        to: target.cellPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableColumnDeletion: {
          tablePosition: target.tablePosition,
          columnIndex: target.leftColumnIndex,
        },
      },
    };
  }

  if (operation.type === "mergeTableCells") {
    const startCell = findEnclosingTableCell(doc, blockFrom);
    if (!startCell) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const table = doc.nodeAt(startCell.tablePosition);
    if (!table || table.type.spec["tableRole"] !== "table") {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const map = TableMap.get(table);
    let rectangle: TableRectangle;
    let endCellPosition: number;
    let endCellEndPosition: number;
    if (operation.rowCount !== undefined) {
      rectangle = {
        left: startCell.leftColumnIndex,
        top: startCell.topRowIndex,
        right: startCell.rightColumnIndex,
        bottom: startCell.topRowIndex + operation.rowCount,
      };
      if (
        rectangle.right - rectangle.left !== 1 ||
        rectangle.bottom > map.height ||
        operation.rowCount < 2
      ) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      const endRelativePosition = map.map[(rectangle.bottom - 1) * map.width + rectangle.left];
      const endCell = endRelativePosition === undefined ? null : table.nodeAt(endRelativePosition);
      if (!endCell || endRelativePosition === undefined) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      endCellPosition = startCell.tablePosition + 1 + endRelativePosition;
      endCellEndPosition = endCellPosition + endCell.nodeSize;
    } else {
      const endTarget = resolveStableBlock({
        snapshot,
        blockId: operation.endBlockId,
        liveBlocks,
        liveBlocksByParaId,
      });
      if (endTarget.type === "skip") {
        return endTarget;
      }
      const endCell = findEnclosingTableCell(doc, endTarget.blockFrom);
      if (!endCell || startCell.tablePosition !== endCell.tablePosition) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      rectangle = {
        left: Math.min(startCell.leftColumnIndex, endCell.leftColumnIndex),
        top: Math.min(startCell.topRowIndex, endCell.topRowIndex),
        right: Math.max(startCell.rightColumnIndex, endCell.rightColumnIndex),
        bottom: Math.max(startCell.bottomRowIndex, endCell.bottomRowIndex),
      };
      endCellPosition = endCell.cellPosition;
      endCellEndPosition = endCell.cellEndPosition;
    }
    if (
      (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) ||
      tableRectangleCutsMergedCell(map, rectangle)
    ) {
      return {
        type: "skip",
        reason:
          rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1
            ? "noopOperation"
            : "unsupportedBlock",
      };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: Math.min(startCell.cellPosition, endCellPosition),
        to: Math.max(startCell.cellEndPosition, endCellEndPosition),
        blockFrom,
        blockTo,
        blockNode,
        tableCellMerge: {
          tablePosition: startCell.tablePosition,
          rectangle,
        },
      },
    };
  }

  if (operation.type === "splitTableCell") {
    const cell = findEnclosingTableCell(doc, blockFrom);
    if (!cell) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const rectangle = {
      left: cell.leftColumnIndex,
      top: cell.topRowIndex,
      right: cell.rightColumnIndex,
      bottom: cell.bottomRowIndex,
    };
    if (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: cell.cellPosition,
        to: cell.cellEndPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableCellSplit: {
          tablePosition: cell.tablePosition,
          rectangle,
        },
      },
    };
  }

  if (operation.type === "deleteBlock" || operation.type === "replaceBlock") {
    const range = getTextRangeFromCleanBlock(cleanBlock);
    if (!range) {
      const insertionPoint = cleanBlock.offsets.at(0);
      if (
        operation.type === "replaceBlock" &&
        currentText.length === 0 &&
        operation.text.length > 0 &&
        insertionPoint !== undefined
      ) {
        return {
          type: "resolved",
          operation: {
            operation,
            from: insertionPoint,
            to: insertionPoint,
            blockFrom,
            blockTo,
            blockNode,
          },
        };
      }
      return { type: "skip", reason: "unsupportedBlock" };
    }
    // The model occasionally emits replaceBlock with text identical
    // to the live block's clean text — usually as a side effect of
    // running through a "review every block" pass. Skip so the
    // panel doesn't show an empty redline (verified in dev-tools
    // trace where one such op surfaced as "Prodávající 3 →
    // Prodávající 3").
    if (operation.type === "replaceBlock" && operation.text === currentText) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: range.from,
        to: range.to,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "replaceInBlock" && operation.find === operation.replace) {
    return { type: "skip", reason: "noopOperation" };
  }

  const quote = getOperationQuote(operation);
  const range = resolveTextInCleanBlock(cleanBlock, quote || currentText);
  if (range.type !== "resolved") {
    return range;
  }

  return {
    type: "resolved",
    operation: {
      operation,
      from: range.from,
      to: range.to,
      blockFrom,
      blockTo,
      blockNode,
    },
  };
};

type OperationResolutionSkip = { type: "skip"; reason: FolioAIEditSkipReason };

const resolveTextInCleanBlock = (
  cleanBlock: { text: string; offsets: number[] },
  find: string,
):
  | { type: "resolved"; from: number; to: number }
  | { type: "skip"; reason: FolioAIEditSkipReason } => {
  if (find.length === 0) {
    return { type: "skip", reason: "emptyOperation" };
  }

  const { text, offsets } = cleanBlock;
  const firstIndex = text.indexOf(find);
  if (firstIndex === -1) {
    return { type: "skip", reason: "missingFind" };
  }
  if (text.includes(find, firstIndex + 1)) {
    return { type: "skip", reason: "ambiguousFind" };
  }

  const from = offsets[firstIndex];
  const to = offsets[firstIndex + find.length];
  if (from === undefined || to === undefined) {
    return { type: "skip", reason: "unsupportedBlock" };
  }

  return { type: "resolved", from, to };
};

const getTextRangeFromCleanBlock = (cleanBlock: {
  text: string;
  offsets: number[];
}): { from: number; to: number } | null => {
  if (cleanBlock.text.length === 0) {
    return null;
  }
  const from = cleanBlock.offsets[0];
  const to = cleanBlock.offsets[cleanBlock.text.length];
  if (from === undefined || to === undefined) {
    return null;
  }
  return { from, to };
};

const getOperationCommentText = (operation: FolioAIEditOperation): string | undefined =>
  "comment" in operation ? operation.comment?.text : undefined;

const getOperationQuote = (operation: FolioAIEditOperation): string | undefined => {
  if (operation.type === "replaceInBlock") {
    return operation.find;
  }
  if (operation.type === "commentOnBlock") {
    return operation.quote;
  }
  return undefined;
};
