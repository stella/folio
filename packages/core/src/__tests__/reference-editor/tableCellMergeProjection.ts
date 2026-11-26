import { Result } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import {
  type DecodedTableCellParagraphSourcePayload,
  decodeTableCellParagraphSourcePayload,
} from "../../docx/paragraphPropertySource";
import { standaloneTableCellFromProseMirror } from "../../prosemirror/conversion/fromProseDoc";
import { standaloneTableCellToProseMirror } from "../../prosemirror/conversion/toProseDoc";
import { getTableCellMergeChange } from "../../prosemirror/tableCellMergeRevision";
import type { TableCell, TableCellFormatting } from "../../types/document";

export const tableCellContinuationPayload = (
  node: PMNode,
): DecodedTableCellParagraphSourcePayload | null => {
  const value = node.attrs["_docxVMergeContinuationCells"];
  return value === undefined || value === null
    ? null
    : decodeTableCellParagraphSourcePayload(value);
};

export const hasMatchingCollapsedTableCellMerge = (
  node: PMNode,
  revisionSet: Set<number> | null,
): boolean => {
  const payload = tableCellContinuationPayload(node);
  return (
    payload !== null &&
    payload.cells.some((cell) => {
      const change = getTableCellMergeChange(cell);
      return (
        change !== null &&
        (revisionSet === null || revisionSet.has(change.info.id))
      );
    })
  );
};

export const tableCellContinuationCells = (
  cell: PMNode,
  rowspan: number,
): TableCell[] => {
  const payload = tableCellContinuationPayload(cell);
  const cells = payload ? [...payload.cells] : [];
  while (cells.length < rowspan - 1) {
    cells.push(emptyVerticalMergeContinuation());
  }
  return cells;
};

export const tableCellContinuationFromNode = (cell: PMNode): TableCell => {
  const continuation = standaloneTableCellFromProseMirror(cell);
  const formatting: TableCellFormatting = {
    ...continuation.formatting,
    vMerge: "continue",
  };
  const colspan = cell.attrs["colspan"];
  if (typeof colspan === "number" && colspan > 1) {
    formatting.gridSpan = colspan;
  }
  continuation.formatting = formatting;
  delete continuation.structuralChange;
  return continuation;
};

const emptyVerticalMergeContinuation = (): TableCell => ({
  type: "tableCell",
  formatting: { vMerge: "continue" },
  content: [{ type: "paragraph", content: [] }],
});

export const createRestoredTableCell = (
  origin: PMNode,
  source: TableCell,
): PMNode | null => {
  const formatting = source.formatting ? { ...source.formatting } : undefined;
  if (formatting) {
    delete formatting.vMerge;
  }
  const restoredSource: TableCell = {
    type: "tableCell",
    ...(formatting ? { formatting } : {}),
    ...(source.propertyChanges
      ? { propertyChanges: source.propertyChanges }
      : {}),
    content: source.content,
  };
  const restored = standaloneTableCellToProseMirror(
    restoredSource,
    origin.type.name === "tableHeader" ? "tableHeader" : "tableCell",
  );
  return Result.try(() =>
    origin.type.schema.nodeFromJSON(restored.toJSON()),
  ).unwrapOr(null);
};
