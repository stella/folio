/** Shared tracked construction for table rows and whole tables. */
import { panic, Result } from "better-result";

import {
  type BlockContent,
  type Document,
  type DocumentBody,
  MAX_REVISION_ID,
  type Paragraph,
  type TableCell,
  type TableRow,
  type TrackedChangeInfo,
} from "../model/document";
import {
  blockListAt,
  endsItsContainer,
  type ParagraphLocation,
  storyParagraphs,
} from "./blocks";
import { structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
  type KeyMembership,
  packageIdentityKeys,
  slotKey,
} from "./ids";
import { childNodes, type InlineNode } from "./leaves";
import { paragraphLength } from "./offsets";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  pendingParagraphMarkReason,
} from "./refusal";
import {
  carriesStamp,
  stampInfo,
  WRAP_KINDS,
  type WrapKind,
  wrapTracked,
} from "./review";
import {
  type DeleteRowOp,
  type DeleteTableOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertRowOp,
  type InsertTableOp,
  type MergeCellsOp,
  type NewIds,
  type RevisionStamp,
} from "./types";

/**
 * A merge tracks the deletion of a row it leaves only continuations, which
 * accepting the merge removes as the direct merge does.
 */
type TrackingOp =
  InsertRowOp | DeleteRowOp | InsertTableOp | DeleteTableOp | MergeCellsOp;
type TrackingRefusalOptions = {
  op: TrackingOp;
  reason: DocumentOpRefusal["reason"];
  message: string;
};
const refused = ({ op, reason, message }: TrackingRefusalOptions) =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));

/** Rebuild the cell's own paragraphs, leaving its nested tables intact. */
const mapParagraphs = (
  blocks: readonly BlockContent[],
  replace: (paragraph: Paragraph) => Paragraph,
): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return replace(block);
      case "blockSdt":
      case "blockCustomXml":
        return { ...block, content: mapParagraphs(block.content, replace) };
      case "table":
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return block;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

const rowWithParagraphs = (
  row: TableRow,
  replace: (paragraph: Paragraph) => Paragraph,
): TableRow => ({
  ...row,
  cells: row.cells.map((cell) => ({
    ...cell,
    content: mapParagraphs(cell.content, replace),
  })),
});

const trackableInline = (nodes: readonly InlineNode[]): boolean =>
  nodes.every((node) => {
    switch (node.type) {
      case "moveFrom":
      case "moveTo":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return false;
      default:
        return trackableInline(childNodes(node) ?? []);
    }
  });

/**
 * A row's structural record carries opaque blocks and nested tables whole.
 * A section break or move requires preserving its other endpoint as well.
 */
const trackableBlocks = (blocks: readonly BlockContent[]): boolean =>
  blocks.every((block) => {
    switch (block.type) {
      case "paragraph":
        return (
          block.sectionProperties === undefined &&
          block.pPrMark?.kind !== "moveFrom" &&
          block.pPrMark?.kind !== "moveTo" &&
          trackableInline(block.content)
        );
      case "blockSdt":
      case "blockCustomXml":
        return trackableBlocks(block.content);
      case "table":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "preservedBlock":
        return true;
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return false;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

const holdsCapturedInline = (items: readonly InlineNode[]): boolean =>
  items.some(
    (item) =>
      item.type === "preservedInline" ||
      item.type === "preservedXml" ||
      holdsCapturedInline(childNodes(item) ?? []),
  );

type TrackTableRowsOptions = {
  document: Document;
  op: TrackingOp;
  rows: readonly TableRow[];
  revision: RevisionStamp;
  newIds: NewIds | undefined;
  kind: WrapKind;
  /**
   * Cells whose content outlives the rows' removal, merged cells the next
   * row continues: their content stays unmarked, as accepting hands it on.
   */
  survivingCells?: ReadonlySet<TableCell>;
};

/** Row stamps own distinct final-cell marks as well as the inline revisions. */
export const trackTableRows = ({
  document,
  op,
  rows,
  revision,
  newIds,
  kind,
  survivingCells = new Set(),
}: TrackTableRowsOptions): Result<TableRow[], DocumentOpRefusal> => {
  if (rows.some((row) => row.structuralChange !== undefined)) {
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message: "A row already carries a structural revision.",
    });
  }
  const finals = new Set<Paragraph>();
  const unmarked = new Set<Paragraph>();
  const before: Paragraph[] = [];
  let missingFinalParagraph = false;
  for (const row of rows) {
    for (const cell of row.cells) {
      const body = { content: cell.content };
      const locations = storyParagraphs(body);
      const handsOn = survivingCells.has(cell);
      for (const { list, paragraph } of locations) {
        before.push(paragraph);
        // The outer row owns the nested table as a block. Leave every
        // inner paragraph and its earlier revisions in the inner scope.
        if (handsOn || list.some((step) => step.kind === "tableCell"))
          unmarked.add(paragraph);
      }
      if (handsOn) continue;
      const final = locations.at(-1);
      if (final === undefined) {
        return refused({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          message: "Every cell must have a paragraph.",
        });
      }
      if (unmarked.has(final.paragraph) || !endsItsContainer(body, final)) {
        missingFinalParagraph = true;
        continue;
      }
      finals.add(final.paragraph);
    }
  }
  const reviewedFinalMark = [...finals].find(
    (paragraph) => paragraph.pPrMark !== undefined,
  )?.pPrMark;
  if (missingFinalParagraph)
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message: "A cell needs a final paragraph to record its mark.",
    });
  if (reviewedFinalMark !== undefined) {
    return refused({
      op,
      reason: pendingParagraphMarkReason(reviewedFinalMark),
      message: "A cell's final paragraph already carries a mark revision.",
    });
  }
  if (
    rows.some((row) =>
      row.cells.some(
        (cell) => !survivingCells.has(cell) && !trackableBlocks(cell.content),
      ),
    )
  ) {
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "The rows contain an unsupported tracked structure.",
    });
  }
  const outside = new Set(packageIdentityKeys(document.package));
  if (
    op.type === DOCUMENT_OP_TYPES.DELETE_ROW ||
    op.type === DOCUMENT_OP_TYPES.DELETE_TABLE ||
    op.type === DOCUMENT_OP_TYPES.MERGE_CELLS
  ) {
    for (const key of identityKeysIn(rows)) outside.delete(key);
  }
  // Reserve every structural id before assigning cell marks or inline ids.
  // Existing row identities remain unavailable to the new structural records.
  const taken = new Set([...outside, ...identityKeysIn(rows)]);
  const firstKey = slotKey({
    space: IDENTITY_SPACES.REVISION,
    id: revision.id,
  });
  outside.add(firstKey);
  taken.add(firstKey);
  const rowStamps: RevisionStamp[] = [];
  const pool = newIds?.revision ?? [];
  let cursor = 0;
  for (const [index] of rows.entries()) {
    if (index === 0) {
      rowStamps.push(revision);
      continue;
    }
    let id = pool[cursor];
    while (
      id !== undefined &&
      Number.isInteger(id) &&
      taken.has(slotKey({ space: IDENTITY_SPACES.REVISION, id }))
    ) {
      cursor += 1;
      id = pool[cursor];
    }
    cursor += 1;
    if (id === undefined) {
      return refused({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message: `The rows need ${rows.length - index} additional structural revision ids.`,
      });
    }
    if (!Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID) {
      return refused({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        message: `${id} cannot be a new revision id.`,
      });
    }
    const key = slotKey({ space: IDENTITY_SPACES.REVISION, id });
    taken.add(key);
    outside.add(key);
    rowStamps.push({ ...revision, id });
  }
  const after: Paragraph[] = [];
  for (const paragraph of before) {
    if (unmarked.has(paragraph)) {
      after.push(paragraph);
      continue;
    }
    // The row owns an opaque paragraph's removal or insertion. Leave its
    // verbatim XML in the original container, outside new inline wrappers.
    const wrapped = holdsCapturedInline(paragraph.content)
      ? { kind: "unchanged" as const }
      : wrapTracked({
          items: paragraph.content,
          from: { offset: 0, zeroWidthBefore: 0 },
          to: {
            offset: paragraphLength(paragraph),
            zeroWidthBefore: Number.MAX_SAFE_INTEGER,
          },
          kind,
          stamp: revision,
          ...(kind === WRAP_KINDS.DELETION
            ? { commentAnchors: "stand" as const }
            : {}),
        });
    if (wrapped.kind === "refused") {
      return refused({
        op,
        reason: wrapped.reason,
        message: "The rows' content cannot record this revision.",
      });
    }
    const content =
      wrapped.kind === "unchanged" ? paragraph.content : wrapped.content;
    if (finals.has(paragraph)) {
      after.push({
        ...paragraph,
        content,
        pPrMark: {
          kind: kind === WRAP_KINDS.INSERTION ? "ins" : "del",
          info: stampInfo(revision),
        },
      });
      continue;
    }
    after.push(
      content === paragraph.content ? paragraph : { ...paragraph, content },
    );
  }
  const freshened = freshenIdentities({
    before,
    after,
    newIds: { ...newIds, revision: pool.slice(cursor) },
    usedElsewhere: () => outside,
  });
  switch (freshened.kind) {
    case "needsIds":
      return refused({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message: `The rows need ${freshened.missing} additional revision ids.`,
      });
    case "invalidId":
      return refused({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        message: `${freshened.id} cannot be a new revision id.`,
      });
    case "fresh": {
      const byId = new Map(
        freshened.paragraphs.map((paragraph) => [
          idKey(paragraph.paraId ?? ""),
          paragraph,
        ]),
      );
      return Result.ok(
        rows.map((row, index) => {
          const stamp =
            rowStamps[index] ??
            panic("A tracked row lost its structural stamp.");
          return {
            ...rowWithParagraphs(
              row,
              (paragraph) =>
                byId.get(idKey(paragraph.paraId ?? "")) ??
                panic("A tracked row lost a paragraph."),
            ),
            structuralChange: {
              type:
                kind === WRAP_KINDS.INSERTION
                  ? "tableRowInsertion"
                  : "tableRowDeletion",
              info: stampInfo(stamp),
            },
          } satisfies TableRow;
        }),
      );
    }
    default: {
      const unreachable: never = freshened;
      return unreachable;
    }
  }
};

/** A cell-final mark is legal only when its innermost row owns its stamp. */
export const permitsCellFinalMark = (
  body: DocumentBody,
  location: ParagraphLocation,
): boolean => {
  const mark = location.paragraph.pPrMark;
  if (mark === undefined || !endsItsContainer(body, location)) return false;
  const stepIndex = location.list.findLastIndex(
    (step) => step.kind === "tableCell",
  );
  const step = location.list[stepIndex];
  if (step?.kind !== "tableCell") return false;
  const table = blockListAt(body.content, location.list.slice(0, stepIndex))[
    step.block
  ];
  const change =
    table?.type === "table"
      ? table.rows[step.row]?.structuralChange
      : undefined;
  if (
    change === undefined ||
    mark.info.id === change.info.id ||
    !structurallyEqual({ ...mark.info, id: 0 }, { ...change.info, id: 0 })
  )
    return false;
  switch (change.type) {
    case "tableRowInsertion":
      return mark.kind === "ins";
    case "tableRowDeletion":
      return mark.kind === "del";
    case "tableCellInsertion":
    case "tableCellDeletion":
    case "tableCellMerge":
      return false;
    default: {
      const unreachable: never = change;
      return unreachable;
    }
  }
};

/**
 * A cell-final mark an exact inverse may write back. Resolving or
 * re-attributing a row's revision apart from its cells' final marks leaves
 * those marks without the row's stamp; the inverse of a later edit to the
 * cells or rows puts them back as they were.
 */
export const permitsRestoredCellFinalMark = ({
  body,
  location,
  opType,
}: {
  body: DocumentBody;
  location: ParagraphLocation;
  opType: DocumentOp["type"];
}): boolean => {
  if (permitsCellFinalMark(body, location)) return true;
  const mark = location.paragraph.pPrMark;
  return (
    (opType === DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS ||
      opType === DOCUMENT_OP_TYPES.SET_TABLE_ROWS) &&
    location.list.some((step) => step.kind === "tableCell") &&
    (mark?.kind === "ins" || mark?.kind === "del")
  );
};

/** Newly stamped table, row, cell and grid ids, including nested tables. */
type StampedTableRevisionIdsOptions = {
  blocks: readonly BlockContent[];
  stamp: RevisionStamp;
  known: KeyMembership;
};
export const stampedTableRevisionIds = ({
  blocks,
  stamp,
  known,
}: StampedTableRevisionIdsOptions): number[] => {
  const out: number[] = [];
  const added = new Set<number>();
  const addId = (id: number) => {
    if (
      known.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })) ||
      added.has(id)
    )
      return;
    added.add(id);
    out.push(id);
  };
  const addInfo = (info: TrackedChangeInfo) => {
    if (carriesStamp(info, stamp)) addId(info.id);
  };
  const visit = (children: readonly BlockContent[]): void => {
    for (const block of children) {
      switch (block.type) {
        case "table":
          for (const change of block.propertyChanges ?? [])
            addInfo(change.info);
          if (block.formatting?.gridChange !== undefined)
            addId(block.formatting.gridChange.id);
          for (const row of block.rows) {
            if (row.structuralChange !== undefined)
              addInfo(row.structuralChange.info);
            for (const change of row.propertyChanges ?? [])
              addInfo(change.info);
            for (const change of row.tablePropertyExceptionChanges ?? [])
              addInfo(change.info);
            for (const cell of row.cells) {
              if (cell.structuralChange !== undefined)
                addInfo(cell.structuralChange.info);
              for (const change of cell.propertyChanges ?? [])
                addInfo(change.info);
              visit(cell.content);
            }
          }
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content);
          break;
        case "paragraph":
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          break;
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    }
  };
  visit(blocks);
  return out;
};
