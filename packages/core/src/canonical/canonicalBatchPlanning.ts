/** Snapshot facts for the shared batch admission and execution owners. */
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import {
  type BatchBlockAccess,
  type BatchOrderAddress,
  contentRank,
  deletionRemovesBlock,
  DOCUMENT_OP_TYPES,
  type EditorIntent,
  type FinalBreakCandidate,
  finalBreakCarrier,
  groupInsertionRuns,
  orderAddressingBeforeRemoval,
  orderBatchOperations,
} from "@stll/docx-core/ops";

import { directJoinTakes } from "../ai-edits/directJoin";
import {
  isInsertedPPrMark,
  snapshotBatchClaim,
  type SnapshotBatchOperation,
} from "../ai-edits/snapshotBatchFacts";
import { findTableColumnInsertion } from "../ai-edits/table-row-column-mutations";
import {
  findEnclosingTableCell,
  findEnclosingTableRow,
} from "../ai-edits/table-targets";
import type { FolioDocumentOperation } from "../document-operations";
import { paragraphEndsItsContainer } from "../prosemirror/containerFinalParagraph";
import type { CanonicalProjection } from "./canonicalSession";

export type CanonicalBatchPlan = {
  claim: ReturnType<typeof snapshotBatchClaim>;
  address: BatchOrderAddress;
  block: number;
  insertionRun: number | null;
};

type CanonicalBatchPlanningOptions = {
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
  from: number;
  to: number;
  projection: CanonicalProjection;
  tracked: boolean;
  inputIndex: number;
};

export const canonicalBatchPlan = ({
  operation,
  intents,
  from,
  to,
  projection,
  tracked,
  inputIndex,
}: CanonicalBatchPlanningOptions): CanonicalBatchPlan => {
  const blockId =
    operation.type === "replaceRange" ||
    operation.type === "formatRange" ||
    operation.type === "commentOnRange"
      ? operation.range.blockId
      : operation.blockId;
  const paragraph =
    projection.paragraph(blockId) ??
    panic("A resolved batch operation lost its source paragraph", {
      id: operation.id,
    });
  const doc = projection.doc;
  const block = paragraph.start - 1;
  const claim = (item: SnapshotBatchOperation) =>
    snapshotBatchClaim({ item, doc, producesTrackedChanges: tracked });
  const content = (
    item: SnapshotBatchOperation,
    position = from,
  ): CanonicalBatchPlan => ({
    claim: claim(item),
    block,
    address: { type: "content", position, rank: 1, inputIndex },
    insertionRun: null,
  });
  const row = () =>
    findEnclosingTableRow(doc, block) ??
    panic("A resolved table operation lost its source row", {
      id: operation.id,
    });
  const cell = () =>
    findEnclosingTableCell(doc, block) ??
    panic("A resolved cell operation lost its source cell", {
      id: operation.id,
    });
  switch (operation.type) {
    case "replaceInBlock":
    case "replaceRange":
      return content({ type: "text", block, from, to });
    case "splitBlock":
      return content({ type: "split", block, from, to });
    case "formatRange":
    case "commentOnRange":
    case "commentOnBlock":
      return {
        claim: claim({ type: "annotation", block, from, to }),
        block,
        address: { type: "annotation", position: from, inputIndex },
        insertionRun: null,
      };
    case "replaceBlock":
      return content(
        intents.some((intent) => intent.type === "replaceText")
          ? { type: "rewriteBlock", block }
          : { type: "paragraphProperties", block },
      );
    case "setBlockParagraphProperties":
      return content({ type: "paragraphProperties", block });
    case "mergeBlockWithNext": {
      const join =
        intents.find((intent) => intent.type === "joinParagraphs") ??
        panic("A resolved merge lost its join intent");
      const next =
        projection.paragraph(join.nextBlockId) ??
        panic("A resolved merge lost its second paragraph");
      return content({
        type: "merge",
        block,
        next: next.start - 1,
        joinsNow: !tracked,
        setsProperties: operation.mergedParagraphProperties !== undefined,
      });
    }
    case "deleteBlock":
      return content({
        type: "deleteBlock",
        block,
        end: block + paragraph.node.nodeSize,
        node: paragraph.node,
      });
    case "insertBeforeBlock":
    case "insertAfterBlock": {
      const plan = content({ type: "insertion", at: from });
      // Matching the native owner: redirected or commented insertions do not
      // share a provisional insertion run with their neighbours.
      const hasComment = operation.comment !== undefined;
      const emptyPageBreak =
        tracked && operation.pageBreakBefore === true && operation.text === "";
      return {
        ...plan,
        insertionRun: hasComment || emptyPageBreak ? null : from,
      };
    }
    case "insertTable":
    case "insertSignatureTable":
      return content({ type: "insertion", at: from });
    case "deleteTable": {
      const target = row();
      return content({
        type: "deleteTable",
        position: target.tablePosition,
        node: target.table,
      });
    }
    case "deleteTableRow": {
      const target = row();
      return content({
        type: "deleteTableRow",
        deletion: target,
        row: target.table.child(target.rowIndex),
        table: target.table,
      });
    }
    case "deleteTableColumn": {
      const target = row();
      const column = cell().leftColumnIndex;
      const plan = content({
        type: "deleteTableColumn",
        deletion: { tablePosition: target.tablePosition, columnIndex: column },
      });
      return {
        ...plan,
        address: {
          type: "column",
          tablePosition: target.tablePosition,
          columnIndex: column,
          insertion: false,
          inputIndex,
        },
      };
    }
    case "insertTableColumn": {
      const insertion =
        findTableColumnInsertion({
          doc,
          blockFrom: block,
          position: operation.position ?? "after",
        }) ?? panic("A resolved column insertion lost its source gap");
      const plan = content({
        type: "tableStructure",
        axis: "column",
        target: {
          type: "tableStructure",
          tablePosition: insertion.tablePosition,
        },
      });
      return {
        ...plan,
        address: {
          type: "column",
          tablePosition: insertion.tablePosition,
          columnIndex: insertion.columnIndex,
          insertion: true,
          inputIndex,
        },
      };
    }
    case "insertTableRow":
      return content({
        type: "tableStructure",
        axis: "row",
        target: { type: "tableStructure", tablePosition: row().tablePosition },
      });
    case "mergeTableCells":
    case "splitTableCell": {
      const target = row();
      const source = cell();
      const merge = intents.find(
        (intent) =>
          intent.type === "table" &&
          intent.operation.type === DOCUMENT_OP_TYPES.MERGE_CELLS,
      );
      const rectangleOfMerge = () => {
        if (
          merge?.type !== "table" ||
          merge.operation.type !== DOCUMENT_OP_TYPES.MERGE_CELLS
        )
          return panic("A resolved cell merge lost its rectangle");
        return merge.operation;
      };
      const rectangle =
        operation.type === "mergeTableCells"
          ? rectangleOfMerge()
          : {
              left: source.leftColumnIndex,
              top: source.topRowIndex,
              right: source.rightColumnIndex,
              bottom: source.bottomRowIndex,
            };
      const plan = content({
        type: "tableStructure",
        axis: "cell",
        target: {
          type:
            operation.type === "mergeTableCells" ? "mergeCells" : "splitCell",
          tablePosition: target.tablePosition,
          rectangle,
        },
      });
      return {
        ...plan,
        address: {
          type: "cell",
          tablePosition: target.tablePosition,
          bottom: rectangle.bottom,
          right: rectangle.right,
          inputIndex,
        },
      };
    }
    default: {
      const unreachable: never = operation;
      return unreachable;
    }
  }
};

type PlannedOperation = {
  operation: FolioDocumentOperation;
  plan: CanonicalBatchPlan;
};

type CanonicalBatchOrderOptions<T extends PlannedOperation> = {
  items: readonly T[];
  doc: PMNode;
};

export const orderCanonicalBatch = <T extends PlannedOperation>({
  items,
  doc,
}: CanonicalBatchOrderOptions<T>): T[] => {
  const deleted = new Set(
    items.flatMap(({ operation, plan }) =>
      operation.type === "deleteBlock" ? [plan.block] : [],
    ),
  );
  const joinedAway = new Set(
    items.flatMap(({ plan }) =>
      plan.claim.type === "merge" && plan.claim.joinsNow ? [plan.block] : [],
    ),
  );
  const deletedSplits = new Set(
    items.flatMap(({ operation, plan }) => {
      if (
        operation.type !== "deleteBlock" ||
        !isInsertedPPrMark(doc.nodeAt(plan.block)?.attrs["pPrMark"])
      )
        return [];
      const at = doc.resolve(plan.block);
      const next = at.parent.maybeChild(at.index() + 1);
      if (next?.type.name !== "paragraph") return [];
      return items.some(
        ({ operation: other, plan: otherPlan }) =>
          other.type === "insertTable" &&
          other.position !== "before" &&
          otherPlan.block === plan.block &&
          otherPlan.claim.type === "insertion" &&
          otherPlan.claim.at ===
            plan.block + (doc.nodeAt(plan.block)?.nodeSize ?? 0),
      )
        ? [plan.block]
        : [];
    }),
  );
  const address = ({ operation, plan }: T): BatchOrderAddress => {
    if (plan.address.type !== "content") return plan.address;
    const hasDeletedSplit = deletedSplits.has(plan.block);
    return {
      ...plan.address,
      position:
        hasDeletedSplit &&
        operation.type === "deleteBlock" &&
        plan.claim.type === "deleteBlock"
          ? plan.claim.end
          : plan.address.position,
      rank: contentRank(operation, { hasDeletedSplit }),
    };
  };
  const access = ({ operation, plan }: T): BatchBlockAccess<number> => {
    const claim = plan.claim;
    if (claim.type === "merge")
      return {
        addresses: [plan.block],
        removes: claim.joinsNow
          ? directJoinTakes(plan.block, claim.next)
          : [claim.next],
      };
    if (operation.type === "deleteTable" || operation.type === "deleteTableRow")
      return { addresses: [], removes: [plan.block] };
    if (plan.address.type === "column" || plan.address.type === "cell")
      return { addresses: [], removes: [] };
    if (operation.type === "insertBeforeBlock" && joinedAway.has(plan.block))
      return { addresses: [], removes: [] };
    if (claim.type !== "deleteBlock")
      return { addresses: [plan.block], removes: [] };
    const at = doc.resolve(plan.block);
    const paragraph =
      doc.nodeAt(plan.block) ??
      panic("A planned deletion lost its source paragraph");
    const before: FinalBreakCandidate<number>[] = [];
    let position = plan.block;
    for (let index = at.index() - 1; index >= 0; index--) {
      const previous = at.parent.child(index);
      position -= previous.nodeSize;
      if (previous.type !== paragraph.type) break;
      before.push({ paragraph: position, deleted: deleted.has(position) });
    }
    const appended = items.some(
      ({ operation: other, plan: otherPlan }) =>
        otherPlan.block === plan.block &&
        (other.type === "insertAfterBlock" ||
          (other.type === "insertTable" && other.position !== "before")),
    );
    const removes = deletionRemovesBlock({
      endsContainer:
        !appended && paragraphEndsItsContainer(at, paragraph.type.name),
      carrier: finalBreakCarrier(before).type,
    });
    return { addresses: [], removes: removes ? [plan.block] : [] };
  };
  const runs = groupInsertionRuns(
    orderBatchOperations(items, address),
    ({ plan }) => plan.insertionRun,
    ({ operation }) => operation.type === "setBlockParagraphProperties",
  );
  return orderAddressingBeforeRemoval(runs, (run) => ({
    addresses: run.flatMap((item) => access(item).addresses),
    removes: run.flatMap((item) => access(item).removes),
  })).flat();
};
