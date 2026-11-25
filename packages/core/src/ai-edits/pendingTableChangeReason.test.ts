/**
 * A tracked table operation over a table, row or cell that already carries a
 * pending change of the kind it would record is refused as
 * `pendingTableChange` on both engines wherever they refuse it, never as a
 * missing capability. Where only one engine refuses the shape, each engine's
 * outcome is stated.
 */
import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditOperation } from "./types";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const AUTHOR = "Reviewer";
const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };
const OWN = { author: AUTHOR, date: "2026-01-01T00:00:00Z" };

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const cell = (paraId: string, text: string): TableCell => ({
  type: "tableCell",
  content: [paragraph(paraId, text)],
});

const row = (cells: readonly TableCell[]): TableRow => ({
  type: "tableRow",
  cells: [...cells],
});

const insertedRow = (
  info: typeof OTHER | typeof OWN,
  cells: readonly TableCell[],
): TableRow => ({
  ...row(cells),
  structuralChange: { type: "tableRowInsertion", info: { id: 6, ...info } },
});

const deletedCell = (paraId: string, text: string, id: number): TableCell => ({
  ...cell(paraId, text),
  structuralChange: { type: "tableCellDeletion", info: { id, ...OTHER } },
});

const body = (table: Table): BlockContent[] => [
  paragraph("00000001", "Before."),
  table,
  paragraph("00000002", "After."),
];

const plainRows = (): TableRow[] => [
  row([cell("00000010", "A1"), cell("00000011", "B1")]),
  row([cell("00000012", ""), cell("00000013", "")]),
];

const STRUCTURAL_REVISION = "A row already carries a structural revision.";

type Outcome = {
  applied: readonly string[];
  skipped: readonly {
    id: string;
    reason: string;
    message?: string;
    canonicalRefusal: undefined;
  }[];
};

const APPLIED: Outcome = { applied: ["op"], skipped: [] };

const refusedAs = (message: string): Outcome => ({
  applied: [],
  skipped: [
    {
      id: "op",
      reason: "pendingTableChange",
      message,
      canonicalRefusal: undefined,
    },
  ],
});

type Shape = {
  table: Table;
  operation: FolioAIEditOperation;
  outcome: Record<HeadlessReviewerSession, Outcome>;
};

const SHAPES = {
  "a row deletion over a row another author's pending insertion added": {
    table: {
      type: "table",
      rows: [
        row([cell("00000010", "A1"), cell("00000011", "B1")]),
        insertedRow(OTHER, [cell("00000012", "A2"), cell("00000013", "B2")]),
      ],
    },
    operation: { id: "op", type: "deleteTableRow", blockId: "00000012" },
    outcome: {
      canonical: refusedAs(STRUCTURAL_REVISION),
    },
  },
  // The author's own pending row is removed outright.
  "a row deletion over the author's own pending row insertion": {
    table: {
      type: "table",
      rows: [
        row([cell("00000010", "A1"), cell("00000011", "B1")]),
        insertedRow(OWN, [cell("00000012", "A2"), cell("00000013", "B2")]),
      ],
    },
    operation: { id: "op", type: "deleteTableRow", blockId: "00000012" },
    outcome: { canonical: APPLIED },
  },
  "a table deletion over a table holding a pending row insertion": {
    table: {
      type: "table",
      rows: [
        row([cell("00000010", "A1"), cell("00000011", "B1")]),
        insertedRow(OTHER, [cell("00000012", "A2"), cell("00000013", "B2")]),
      ],
    },
    operation: { id: "op", type: "deleteTable", blockId: "00000010" },
    outcome: {
      canonical: refusedAs(STRUCTURAL_REVISION),
    },
  },
  "a column deletion over cells already suggested for deletion": {
    table: {
      type: "table",
      rows: [
        row([deletedCell("00000010", "A1", 6), cell("00000011", "B1")]),
        row([deletedCell("00000012", "A2", 7), cell("00000013", "B2")]),
      ],
    },
    operation: { id: "op", type: "deleteTableColumn", blockId: "00000010" },
    outcome: {
      canonical: refusedAs(
        "The column's cells are already suggested for deletion.",
      ),
    },
  },
  // The table keeps one pending change per element.
  "a column insertion in a table with a pending property change": {
    table: {
      type: "table",
      propertyChanges: [
        {
          type: "tablePropertyChange",
          info: { id: 9, ...OTHER },
          previousFormatting: { layout: "fixed" },
        },
      ],
      rows: plainRows(),
    },
    operation: {
      id: "op",
      type: "insertTableColumn",
      blockId: "00000010",
      position: "after",
      cellTexts: ["New"],
    },
    outcome: {
      canonical: refusedAs(
        "Resolve table reviews before changing its columns.",
      ),
    },
  },
  "a cell merge in a table with a pending property change": {
    table: {
      type: "table",
      propertyChanges: [
        {
          type: "tablePropertyChange",
          info: { id: 9, ...OTHER },
          previousFormatting: { layout: "fixed" },
        },
      ],
      rows: plainRows(),
    },
    operation: {
      id: "op",
      type: "mergeTableCells",
      blockId: "00000010",
      rowCount: 2,
    },
    outcome: {
      canonical: refusedAs("Resolve table reviews before changing cells."),
    },
  },
} as const satisfies Record<string, Shape>;

const source = async (table: Table) => {
  const document: Document = createEmptyDocument();
  document.package.document.content = body(table);
  return await createDocx(document);
};

describe.each(SESSIONS)("%s: a tracked table operation", (session) => {
  test.each(Object.entries(SHAPES))(
    "%s is applied, or refused as pendingTableChange",
    async (_, { table, operation, outcome }) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await source(table), {
        author: AUTHOR,
        session,
      });
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [operation],
        },
        { undefinedReferences: "refuse" },
      );
      expect({
        applied: result.applied.map(({ id }) => id),
        skipped: result.skipped.map(
          ({ id, reason, message, canonicalRefusal }) => ({
            id,
            reason,
            message,
            canonicalRefusal,
          }),
        ),
      }).toEqual({
        applied: [...outcome[session].applied],
        skipped: [...outcome[session].skipped],
      });
    },
  );
});
