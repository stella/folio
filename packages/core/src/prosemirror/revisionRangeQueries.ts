import type { EditorState } from "prosemirror-state";

import {
  expectRunPropertyChangeMarkAttrs,
  expectTrackedChangeMarkAttrs,
} from "./attrs";
import { enclosingRevisionIds } from "./contentControlRevisions";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "./extensions/nodes/SdtExtension";
import { getFolioNodeRevisionCarriers } from "./revisionCarriers";
import { getTableCellMergeChange } from "./tableCellMergeRevision";

export type TableRowRevisionAttr = {
  revisionId: number;
  author?: string;
  date?: string | null;
};

type TableCellRevisionAttr =
  | {
      kind: "ins" | "del";
      info: {
        revisionId: number;
      };
    }
  | {
      kind: "merge";
      info: {
        revisionId: number;
      };
      verticalMerge?: "continue" | "rest";
      verticalMergeOriginal?: "continue" | "rest";
    };

export function isTableRowRevisionAttr(
  value: unknown,
): value is TableRowRevisionAttr {
  return (
    typeof value === "object" &&
    value !== null &&
    "revisionId" in value &&
    typeof value.revisionId === "number"
  );
}

export function isTableCellRevisionAttr(
  value: unknown,
): value is TableCellRevisionAttr {
  if (
    typeof value !== "object" ||
    value === null ||
    !("kind" in value) ||
    !("info" in value)
  ) {
    return false;
  }
  if (value.kind !== "ins" && value.kind !== "del" && value.kind !== "merge") {
    return false;
  }
  const info = value.info;
  return (
    typeof info === "object" &&
    info !== null &&
    "revisionId" in info &&
    typeof info.revisionId === "number"
  );
}

/**
 * Find the document range covered by all revision carriers with any of the
 * given ids. Returns null when none are present (already accepted/rejected, or
 * never existed). A replace operation typically passes two ids; standalone
 * text, formatting, paragraph, section, and table revisions pass a single id.
 */
export function findAIEditRevisionRange(
  state: EditorState,
  revisionIds: number | readonly number[],
): { from: number; to: number } | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  const runPropertyChangeType = state.schema.marks["runPropertyChange"];
  const idSet = new Set<number>(
    typeof revisionIds === "number" ? [revisionIds] : revisionIds,
  );

  const range = { from: null as number | null, to: null as number | null };
  const includeRange = (from: number, to: number): void => {
    if (range.from === null || from < range.from) {
      range.from = from;
    }
    if (range.to === null || to > range.to) {
      range.to = to;
    }
  };

  state.doc.descendants((node, pos) => {
    for (const carrier of getFolioNodeRevisionCarriers(node, pos)) {
      if (idSet.has(carrier.id)) {
        includeRange(carrier.from, carrier.to);
      }
    }
    // A control an id encloses goes with it, so the range spans the control.
    if (
      node.type.name === INLINE_CONTENT_CONTROL_NODE_NAME &&
      enclosingRevisionIds(node).some((id) => idSet.has(id))
    ) {
      includeRange(pos, pos + node.nodeSize);
    }
    if (node.type.name === "tableRow") {
      for (const attrName of ["trIns", "trDel"] as const) {
        const marker = node.attrs[attrName];
        if (isTableRowRevisionAttr(marker) && idSet.has(marker.revisionId)) {
          includeRange(pos, pos + node.nodeSize);
          return false;
        }
      }
    }
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const marker = node.attrs["cellMarker"];
      const hasDirectRevision =
        isTableCellRevisionAttr(marker) && idSet.has(marker.info.revisionId);
      const continuationCells = node.attrs["_docxVMergeContinuationCells"];
      const hasCollapsedRevision =
        Array.isArray(continuationCells) &&
        continuationCells.some((cell) => {
          const change = getTableCellMergeChange(cell);
          return change !== null && idSet.has(change.info.id);
        });
      if (hasDirectRevision || hasCollapsedRevision) {
        includeRange(pos, pos + node.nodeSize);
        return false;
      }
    }
    // Widen from `isText` to `isInline` so an AI-edit revision on an inline
    // atom (image, shape) shows up in the matched range. eigenpal #641.
    if (!node.isInline) {
      return;
    }
    for (const mark of node.marks) {
      if (
        mark.type === runPropertyChangeType &&
        expectRunPropertyChangeMarkAttrs(mark).changes.some((change) =>
          idSet.has(change.info.id),
        )
      ) {
        includeRange(pos, pos + node.nodeSize);
        break;
      }
      if (
        (mark.type === insertionType || mark.type === deletionType) &&
        (idSet.has(expectTrackedChangeMarkAttrs(mark).revisionId) ||
          (
            expectTrackedChangeMarkAttrs(mark)._docxRevisionAncestors ?? []
          ).some((ancestor) => idSet.has(ancestor.revisionId)))
      ) {
        includeRange(pos, pos + node.nodeSize);
        break;
      }
    }
    return undefined;
  });

  if (range.from === null || range.to === null) {
    return null;
  }
  return { from: range.from, to: range.to };
}
