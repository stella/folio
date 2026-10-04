/** Shared tracked construction for table rows and whole tables. */
import { Result, panic } from "better-result";

import {
  MAX_REVISION_ID,
  type BlockContent,
  type Document,
  type DocumentBody,
  type Paragraph,
  type TableRow,
  type TrackedChangeInfo,
} from "../model/document";
import { blockListAt, endsItsContainer, storyParagraphs, type ParagraphLocation } from "./blocks";
import { structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import { IDENTITY_SPACES, identityKeysIn, idKey, packageIdentityKeys, slotKey } from "./ids";
import { childNodes, type InlineNode } from "./leaves";
import { paragraphLength } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { carriesStamp, stampInfo, wrapTracked, WRAP_KINDS, type WrapKind } from "./review";
import {
  DOCUMENT_OP_TYPES,
  type DeleteRowOp,
  type DeleteTableOp,
  type InsertRowOp,
  type InsertTableOp,
  type NewIds,
  type RevisionStamp,
} from "./types";

type TrackingOp = InsertRowOp | DeleteRowOp | InsertTableOp | DeleteTableOp;
type TrackingRefusalOptions = {
  op: TrackingOp;
  reason: DocumentOpRefusal["reason"];
  message: string;
};
const refused = ({ op, reason, message }: TrackingRefusalOptions) =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));

const paragraphsIn = (rows: readonly TableRow[]): Paragraph[] =>
  storyParagraphs({ content: [{ type: "table", rows: [...rows] }] }).map(
    ({ paragraph }) => paragraph,
  );

/** Rebuild only paths to changed paragraphs, retaining every other record. */
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
        return {
          ...block,
          rows: block.rows.map((row) => rowWithParagraphs(row, replace)),
        };
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
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
      case "preservedInline":
      case "preservedXml":
        return false;
      default:
        return trackableInline(childNodes(node) ?? []);
    }
  });

/** Nested tables, captured blocks and moves have no tracked-row construction. */
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
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return false;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

type TrackTableRowsOptions = {
  document: Document;
  op: TrackingOp;
  rows: readonly TableRow[];
  revision: RevisionStamp;
  newIds: NewIds | undefined;
  kind: WrapKind;
};

/** Row stamps own distinct final-cell marks as well as the inline revisions. */
export const trackTableRows = ({
  document,
  op,
  rows,
  revision,
  newIds,
  kind,
}: TrackTableRowsOptions): Result<TableRow[], DocumentOpRefusal> => {
  if (rows.some((row) => row.structuralChange !== undefined)) {
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "A row already carries a structural revision.",
    });
  }
  const finals = new Set<Paragraph>();
  let missingFinalParagraph = false;
  for (const row of rows) {
    for (const cell of row.cells) {
      const body = { content: cell.content };
      const final = storyParagraphs(body).at(-1);
      if (final === undefined) {
        return refused({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          message: "Every cell must have a paragraph.",
        });
      }
      if (!endsItsContainer(body, final)) {
        missingFinalParagraph = true;
        continue;
      }
      finals.add(final.paragraph);
    }
  }
  if (
    missingFinalParagraph ||
    rows.some(
      (row) =>
        row.preserved !== undefined ||
        row.bookmarks !== undefined ||
        row.contentControls !== undefined ||
        row.carrierStack !== undefined ||
        row.cells.some(
          (cell) =>
            cell.structuralChange !== undefined ||
            (cell.propertyChanges?.length ?? 0) > 0 ||
            cell.contentControls !== undefined ||
            cell.carrierStack !== undefined ||
            !trackableBlocks(cell.content),
        ),
    )
  ) {
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "The rows contain an unsupported tracked structure.",
    });
  }
  if ([...finals].some((paragraph) => paragraph.pPrMark !== undefined)) {
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "A cell's final paragraph already carries a mark revision.",
    });
  }
  const before = paragraphsIn(rows);
  const outside = new Set(packageIdentityKeys(document.package));
  if (op.type === DOCUMENT_OP_TYPES.DELETE_ROW || op.type === DOCUMENT_OP_TYPES.DELETE_TABLE) {
    for (const key of identityKeysIn(rows)) outside.delete(key);
  }
  // Reserve every structural id before assigning cell marks or inline ids.
  // Existing row identities remain unavailable to the new structural records.
  const taken = new Set([...outside, ...identityKeysIn(rows)]);
  const firstKey = slotKey({ space: IDENTITY_SPACES.REVISION, id: revision.id });
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
    const wrapped = wrapTracked({
      items: paragraph.content,
      from: { offset: 0, zeroWidthBefore: 0 },
      to: { offset: paragraphLength(paragraph), zeroWidthBefore: Number.MAX_SAFE_INTEGER },
      kind,
      stamp: revision,
    });
    if (wrapped.kind === "refused") {
      return refused({
        op,
        reason: wrapped.reason,
        message: "The rows' content cannot record this revision.",
      });
    }
    const content = wrapped.kind === "unchanged" ? paragraph.content : wrapped.content;
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
    after.push(content === paragraph.content ? paragraph : { ...paragraph, content });
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
        freshened.paragraphs.map((paragraph) => [idKey(paragraph.paraId ?? ""), paragraph]),
      );
      return Result.ok(
        rows.map((row, index) => {
          const stamp = rowStamps[index] ?? panic("A tracked row lost its structural stamp.");
          return {
            ...rowWithParagraphs(
              row,
              (paragraph) =>
                byId.get(idKey(paragraph.paraId ?? "")) ?? panic("A tracked row lost a paragraph."),
            ),
            structuralChange: {
              type: kind === WRAP_KINDS.INSERTION ? "tableRowInsertion" : "tableRowDeletion",
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
export const permitsCellFinalMark = (body: DocumentBody, location: ParagraphLocation): boolean => {
  const mark = location.paragraph.pPrMark;
  if (mark === undefined || !endsItsContainer(body, location)) return false;
  const stepIndex = location.list.findLastIndex((step) => step.kind === "tableCell");
  const step = location.list[stepIndex];
  if (step?.kind !== "tableCell") return false;
  const table = blockListAt(body.content, location.list.slice(0, stepIndex))[step.block];
  const change = table?.type === "table" ? table.rows[step.row]?.structuralChange : undefined;
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

/** Newly stamped table, row, cell and grid ids, including nested tables. */
type StampedTableRevisionIdsOptions = {
  blocks: readonly BlockContent[];
  stamp: RevisionStamp;
  known: ReadonlySet<string>;
};
export const stampedTableRevisionIds = ({
  blocks,
  stamp,
  known,
}: StampedTableRevisionIdsOptions): number[] => {
  const out: number[] = [];
  const added = new Set<number>();
  const addId = (id: number) => {
    if (known.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })) || added.has(id)) return;
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
          for (const change of block.propertyChanges ?? []) addInfo(change.info);
          if (block.formatting?.gridChange !== undefined) addId(block.formatting.gridChange.id);
          for (const row of block.rows) {
            if (row.structuralChange !== undefined) addInfo(row.structuralChange.info);
            for (const change of row.propertyChanges ?? []) addInfo(change.info);
            for (const change of row.tablePropertyExceptionChanges ?? []) addInfo(change.info);
            for (const cell of row.cells) {
              if (cell.structuralChange !== undefined) addInfo(cell.structuralChange.info);
              for (const change of cell.propertyChanges ?? []) addInfo(change.info);
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
