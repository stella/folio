import type { TableBlockIntentOperation } from "../../src/ops/editorIntent";
import {
  type BlockInsertionPoint,
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  OP_STORIES,
} from "../../src/ops/types";

const base = {
  type: DOCUMENT_OP_TYPES.INSERT_TABLE,
  story: OP_STORIES.MAIN,
  blockId: "12345678",
  table: {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [{ type: "paragraph", paraId: "23456789", content: [] }],
          },
        ],
      },
    ],
  },
} as const satisfies TableBlockIntentOperation;

const implicitParagraph = {
  ...base,
  separated: true,
  paragraphAfter: "none",
} as const satisfies TableBlockIntentOperation;
const explicitParagraph = {
  ...base,
  anchor: "paragraph",
  paragraphAfter: "empty",
} as const satisfies TableBlockIntentOperation;
const tableAnchor = {
  ...base,
  anchor: "table",
} as const satisfies TableBlockIntentOperation;

const separatedInput = {
  ...tableAnchor,
  separated: true,
} as const;
const trailingInput = {
  ...tableAnchor,
  paragraphAfter: "none",
} as const;
// @ts-expect-error A table anchor cannot request separating paragraphs.
const separatedTable: TableBlockIntentOperation = separatedInput;
// @ts-expect-error A table anchor cannot request a trailing paragraph.
const trailingTable: TableBlockIntentOperation = trailingInput;

const listStart = {
  ...base,
  anchor: "containerStart",
} as const satisfies TableBlockIntentOperation;
const sidedListStartInput = { ...listStart, side: "before" } as const;
const separatedListStartInput = { ...listStart, separated: true } as const;
// @ts-expect-error A list start is one place: it takes no side.
const sidedListStart: TableBlockIntentOperation = sidedListStartInput;
// @ts-expect-error A list start cannot request separating paragraphs.
const separatedListStart: TableBlockIntentOperation = separatedListStartInput;

void implicitParagraph;
void explicitParagraph;
void tableAnchor;
void separatedTable;
void trailingTable;
void listStart;
void sidedListStart;
void separatedListStart;

const tableSideParagraphs = {
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "beforeTable", blockId: "23456789" },
  blocks: [{ type: "paragraph", paraId: "34567890", content: [] }],
} as const satisfies InsertBlocksOp;
const paragraphSide = {
  type: "before",
  blockId: "23456789",
} as const satisfies BlockInsertionPoint;
// @ts-expect-error Paragraph-only insertion points cannot stand in for a table boundary.
const tableAsParagraph: BlockInsertionPoint = tableSideParagraphs.at;
const unknownBoundary = {
  ...tableSideParagraphs,
  at: { type: "insideTable", blockId: "23456789" },
} as const;
// @ts-expect-error Table insertion is beside its table, never an implicit cell insertion.
const unsupportedBoundary: InsertBlocksOp = unknownBoundary;
void tableSideParagraphs;
void paragraphSide;
void tableAsParagraph;
void unsupportedBoundary;
