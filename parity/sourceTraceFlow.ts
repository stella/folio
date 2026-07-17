import type {
  FlowBlock,
  ParagraphBlock,
  Run,
  TableBlock,
  TableCell,
} from "../packages/core/src/layout-engine/types";
import { normalizeLineText } from "./textNorm";

type TracePathSegment = number | "blocks" | "cells" | "content" | "rows";

type FlowBlockSummary =
  | {
      type: "paragraph";
      id: ParagraphBlock["id"];
      text: string;
      attrs: ParagraphBlock["attrs"];
      runs: unknown[];
    }
  | {
      type: "table";
      id: TableBlock["id"];
      rowCount: number;
    }
  | {
      type: "block";
      id: FlowBlock["id"];
      kind: Exclude<FlowBlock["kind"], "paragraph" | "table">;
    };

type TableCellSummary = {
  cellIndex: number;
  cellId: TableCell["id"];
  text: string;
  blocks: FlowBlockSummary[];
};

type ParagraphFlowMatch = {
  type: "paragraph";
  path: TracePathSegment[];
  id: ParagraphBlock["id"];
  text: string;
  attrs: ParagraphBlock["attrs"];
  runs: unknown[];
};

type TableRowFlowMatch = {
  type: "tableRow";
  path: TracePathSegment[];
  tableId: TableBlock["id"];
  rowId: TableBlock["rows"][number]["id"];
  rowIndex: number;
  text: string;
  cells: TableCellSummary[];
};

export type FlowMatch = ParagraphFlowMatch | TableRowFlowMatch;

const truncate = (value: string, max = 240): string =>
  value.length <= max ? value : `${value.slice(0, max - 1)}…`;

const normalizedSearchText = (value: string): string =>
  normalizeLineText(value).toLocaleLowerCase();

const summarizeRun = (run: Run): unknown => {
  const base = {
    kind: run.kind,
    pmStart: run.pmStart,
    pmEnd: run.pmEnd,
    fontFamily: "fontFamily" in run ? run.fontFamily : undefined,
    fontSize: "fontSize" in run ? run.fontSize : undefined,
    bold: "bold" in run ? run.bold : undefined,
    allCaps: "allCaps" in run ? run.allCaps : undefined,
    smallCaps: "smallCaps" in run ? run.smallCaps : undefined,
  };
  if (run.kind === "text") {
    return { ...base, text: truncate(run.text) };
  }
  if (run.kind === "field") {
    return {
      ...base,
      fieldType: run.fieldType,
      instruction: run.instruction,
      fallback: run.fallback,
      fldLock: run.fldLock,
    };
  }
  if (run.kind === "image") {
    return { ...base, width: run.width, height: run.height, wrapType: run.wrapType };
  }
  if (run.kind === "math") {
    return { ...base, plainText: run.plainText };
  }
  return base;
};

const paragraphText = (block: ParagraphBlock): string =>
  block.runs
    .map((run) => {
      if (run.kind === "text") return run.text;
      if (run.kind === "field") return run.fallback ?? "";
      if (run.kind === "tab") return "\t";
      return "";
    })
    .join("");

const blockText = (block: FlowBlock): string => {
  if (block.kind === "paragraph") {
    return paragraphText(block);
  }
  if (block.kind === "table") {
    return block.rows
      .flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .map(blockText)
      .join(" ");
  }
  if (block.kind === "textBox") {
    return block.content.map(blockText).join(" ");
  }
  return "";
};

const summarizeBlock = (block: FlowBlock): FlowBlockSummary => {
  if (block.kind === "paragraph") {
    return {
      type: "paragraph",
      id: block.id,
      text: truncate(paragraphText(block)),
      attrs: block.attrs,
      runs: block.runs.map(summarizeRun),
    };
  }
  if (block.kind === "table") {
    return { type: "table", id: block.id, rowCount: block.rows.length };
  }
  return { type: "block", id: block.id, kind: block.kind };
};

const summarizeCell = (cell: TableCell, cellIndex: number): TableCellSummary => ({
  cellIndex,
  cellId: cell.id,
  text: truncate(cell.blocks.map(blockText).join(" ")),
  blocks: cell.blocks.map(summarizeBlock),
});

type VisitBlocksOptions = {
  blocks: FlowBlock[];
  path: TracePathSegment[];
  needle: string;
  limit: number;
  matches: FlowMatch[];
};

type VisitTableOptions = Omit<VisitBlocksOptions, "blocks"> & {
  table: TableBlock;
};

const visitBlocks = ({ blocks, path, needle, limit, matches }: VisitBlocksOptions): void => {
  for (const [blockIndex, block] of blocks.entries()) {
    if (matches.length >= limit) {
      return;
    }

    const blockPath = [...path, blockIndex];
    if (block.kind === "paragraph") {
      const text = paragraphText(block);
      if (normalizedSearchText(text).includes(needle)) {
        matches.push({
          type: "paragraph",
          path: blockPath,
          id: block.id,
          text: truncate(text),
          attrs: block.attrs,
          runs: block.runs.map(summarizeRun),
        });
      }
      continue;
    }

    if (block.kind === "table") {
      visitTable({ table: block, path: blockPath, needle, limit, matches });
      continue;
    }

    if (block.kind === "textBox") {
      visitBlocks({
        blocks: block.content,
        path: [...blockPath, "content"],
        needle,
        limit,
        matches,
      });
    }
  }
};

const visitTable = ({ table, path, needle, limit, matches }: VisitTableOptions): void => {
  for (const [rowIndex, row] of table.rows.entries()) {
    if (matches.length >= limit) {
      return;
    }

    const text = row.cells.map((cell) => cell.blocks.map(blockText).join(" ")).join(" ");
    if (normalizedSearchText(text).includes(needle)) {
      matches.push({
        type: "tableRow",
        path: [...path, "rows", rowIndex],
        tableId: table.id,
        rowId: row.id,
        rowIndex,
        text: truncate(text),
        cells: row.cells.map(summarizeCell),
      });
    }

    for (const [cellIndex, cell] of row.cells.entries()) {
      visitBlocks({
        blocks: cell.blocks,
        path: [...path, "rows", rowIndex, "cells", cellIndex, "blocks"],
        needle,
        limit,
        matches,
      });
    }
  }
};

type TraceFlowBlocksOptions = {
  blocks: FlowBlock[];
  query: string;
  limit: number;
};

export const traceFlowBlocks = ({ blocks, query, limit }: TraceFlowBlocksOptions): FlowMatch[] => {
  const matches: FlowMatch[] = [];
  visitBlocks({
    blocks,
    path: ["blocks"],
    needle: normalizedSearchText(query),
    limit,
    matches,
  });
  return matches;
};
