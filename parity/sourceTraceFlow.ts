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

type BlockTextReader = (block: FlowBlock) => string;

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

const createBlockTextReader = (): BlockTextReader => {
  const cache = new WeakMap<FlowBlock, string>();
  const blockText: BlockTextReader = (block) => {
    const cached = cache.get(block);
    if (cached !== undefined) {
      return cached;
    }

    let text = "";
    if (block.kind === "paragraph") {
      text = paragraphText(block);
    } else if (block.kind === "table") {
      text = block.rows
        .flatMap((row) => row.cells)
        .flatMap((cell) => cell.blocks)
        .map(blockText)
        .join(" ");
    } else if (block.kind === "textBox") {
      text = block.content.map(blockText).join(" ");
    }
    cache.set(block, text);
    return text;
  };
  return blockText;
};

const summarizeBlock = (block: FlowBlock, blockText: BlockTextReader): FlowBlockSummary => {
  if (block.kind === "paragraph") {
    return {
      type: "paragraph",
      id: block.id,
      text: truncate(blockText(block)),
      attrs: block.attrs,
      runs: block.runs.map(summarizeRun),
    };
  }
  if (block.kind === "table") {
    return { type: "table", id: block.id, rowCount: block.rows.length };
  }
  return { type: "block", id: block.id, kind: block.kind };
};

type SummarizeCellOptions = {
  cell: TableCell;
  cellIndex: number;
  blockText: BlockTextReader;
};

const summarizeCell = ({ cell, cellIndex, blockText }: SummarizeCellOptions): TableCellSummary => ({
  cellIndex,
  cellId: cell.id,
  text: truncate(cell.blocks.map(blockText).join(" ")),
  blocks: cell.blocks.map((block) => summarizeBlock(block, blockText)),
});

type TraceContext = {
  needle: string;
  limit: number;
  matches: FlowMatch[];
  blockText: BlockTextReader;
};

type VisitBlocksOptions = {
  blocks: FlowBlock[];
  path: TracePathSegment[];
  context: TraceContext;
};

type VisitTableOptions = Omit<VisitBlocksOptions, "blocks"> & {
  table: TableBlock;
};

const visitBlocks = ({ blocks, path, context }: VisitBlocksOptions): void => {
  for (const [blockIndex, block] of blocks.entries()) {
    if (context.matches.length >= context.limit) {
      return;
    }

    const blockPath = [...path, blockIndex];
    if (block.kind === "paragraph") {
      const text = context.blockText(block);
      if (normalizedSearchText(text).includes(context.needle)) {
        context.matches.push({
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
      visitTable({ table: block, path: blockPath, context });
      continue;
    }

    if (block.kind === "textBox") {
      visitBlocks({
        blocks: block.content,
        path: [...blockPath, "content"],
        context,
      });
    }
  }
};

const visitTable = ({ table, path, context }: VisitTableOptions): void => {
  for (const [rowIndex, row] of table.rows.entries()) {
    if (context.matches.length >= context.limit) {
      return;
    }

    const text = row.cells.map((cell) => cell.blocks.map(context.blockText).join(" ")).join(" ");
    if (normalizedSearchText(text).includes(context.needle)) {
      context.matches.push({
        type: "tableRow",
        path: [...path, "rows", rowIndex],
        tableId: table.id,
        rowId: row.id,
        rowIndex,
        text: truncate(text),
        cells: row.cells.map((cell, cellIndex) =>
          summarizeCell({ cell, cellIndex, blockText: context.blockText }),
        ),
      });
    }

    for (const [cellIndex, cell] of row.cells.entries()) {
      visitBlocks({
        blocks: cell.blocks,
        path: [...path, "rows", rowIndex, "cells", cellIndex, "blocks"],
        context,
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
    context: {
      needle: normalizedSearchText(query),
      limit,
      matches,
      blockText: createBlockTextReader(),
    },
  });
  return matches;
};
