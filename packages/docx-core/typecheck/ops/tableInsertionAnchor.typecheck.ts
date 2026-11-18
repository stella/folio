import type { TableBlockIntentOperation } from "../../src/ops/editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../../src/ops/types";

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

void implicitParagraph;
void explicitParagraph;
void tableAnchor;
void separatedTable;
void trailingTable;
