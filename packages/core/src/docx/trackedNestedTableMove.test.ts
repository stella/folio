import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import type { Document, Paragraph, Table, TableCell } from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const paragraph = (paraId: string, text?: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === undefined
      ? []
      : [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (...content: TableCell["content"]): TableCell => ({
  type: "tableCell",
  content,
});
const nested: Table = {
  type: "table",
  columnWidths: [900],
  rows: [
    { type: "tableRow", cells: [cell(paragraph("00000003", "Nested 1"))] },
    { type: "tableRow", cells: [cell(paragraph("00000004", "Nested 2"))] },
  ],
};

const tableOf = (document: Document): Table => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "table") throw new TypeError("Expected a table.");
  return block;
};

const projection = (table: Table): unknown =>
  table.rows.map((row) =>
    row.cells.map((item) => ({
      span: item.formatting?.gridSpan ?? 1,
      merge: item.formatting?.vMerge,
      content: item.content.map((block) => {
        if (block.type === "table") return projection(block);
        if (block.type === "paragraph")
          return { paraId: block.paraId, content: block.content };
        return block;
      }),
    })),
  );

test.each(["horizontal", "rectangular", "retained"] as const)(
  "a %s merge moves a nested table and closing paragraph through native review",
  (shape) => {
    const rectangular = shape !== "horizontal";
    const retained = shape === "retained";
    const owner = cell(paragraph("00000001", retained ? undefined : "Owner"));
    const other = cell(
      paragraph("00000002", "Before"),
      nested,
      paragraph("00000005"),
    );
    const first = {
      type: "tableRow",
      cells: [
        owner,
        rectangular ? cell(paragraph("00000006")) : other,
        ...(retained ? [cell(paragraph("00000009", "Outside"))] : []),
      ],
    } satisfies Table["rows"][number];
    const table: Table = {
      type: "table",
      columnWidths: [1000, 1400, ...(retained ? [1600] : [])],
      rows: [
        first,
        ...(rectangular
          ? [
              {
                type: "tableRow" as const,
                cells: [
                  cell(paragraph("00000007")),
                  other,
                  ...(retained ? [cell(paragraph("0000000A", "Outside"))] : []),
                ],
              },
            ]
          : []),
      ],
    };
    const ending = paragraph("00000008");
    const withTable = (block: Table): Document => ({
      package: { document: { content: [block, ending] } },
    });
    const original = withTable(table);
    const operation = {
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      top: 0,
      bottom: rectangular ? 2 : 1,
      left: 0,
      right: 2,
    } as const;
    const intent = { type: "table", operation } as const;
    const ids = createEditorIntentIdAllocator()(original, intent);
    const compiled = compileEditorIntent(original, {
      intent,
      mode: {
        type: "suggesting",
        revision: {
          id: ids.revisionId,
          author: "Reviewer",
          date: "2026-11-27T00:00:00Z",
        },
        newIds: ids.newIds,
      },
    }).unwrap();
    const op = compiled.ops.at(0);
    if (op?.type !== DOCUMENT_OP_TYPES.MERGE_CELLS)
      throw new TypeError("Expected a merge operation.");
    const continuations = Number(rectangular);
    expect(op.newBlockIds).toHaveLength(continuations + 4);
    const direct = applyDocumentOp(original, {
      ...operation,
      newBlockIds: op.newBlockIds.slice(0, continuations),
    }).unwrap().document;
    const pending = applyDocumentOps(original, compiled.ops).unwrap();
    expect(
      applyDocumentOps(pending.document, pending.inverse).unwrap().document,
    ).toStrictEqual(original);

    const marked = tableOf(pending.document);
    const geometry = marked.formatting?.gridChange?.geometryChange;
    const nestedId =
      geometry?.type === "mergeCells"
        ? geometry.nestedMoves?.at(0)?.rows.at(0)
        : undefined;
    if (nestedId === undefined) throw new Error("Expected a nested row stamp.");
    const xml = serializeTable(marked, serializeParagraph);
    expect(xml).toContain(`<w:trPr><w:ins w:id="${nestedId}"`);
    expect(xml.match(/<w:tbl>/gu)).toHaveLength(3);
    const element = parseXmlDocument(
      xml.replace(
        "<w:tbl",
        `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
      ),
    );
    if (element === null) throw new Error("Expected well-formed XML.");
    const parsed = parseTable(element, new Map(), null, null, null, null);
    if (parsed === undefined) throw new Error("Expected a reopened table.");

    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const expected =
        decision === REVISION_DECISIONS.ACCEPT ? direct : original;
      const resolve = (document: Document) =>
        applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [nestedId],
          decision,
        }).unwrap();
      const result = resolve(pending.document);
      expect(result.document).toStrictEqual(expected);
      expect(
        applyDocumentOps(result.document, result.inverse).unwrap().document,
      ).toStrictEqual(pending.document);
      expect(
        projection(tableOf(resolve(withTable(parsed)).document)),
      ).toStrictEqual(projection(tableOf(expected)));
    }
  },
);

test("a merge refuses before creating an unwriteable nested-table review record", () => {
  const destinationId = (index: number) =>
    (0x4000 + index).toString(16).padStart(8, "0");
  const nestedTable: Table = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          cell(
            ...Array.from({ length: 1600 }, (_, index) =>
              paragraph(destinationId(index)),
            ),
          ),
        ],
      },
    ],
  };
  const original = {
    package: {
      document: {
        content: [
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  cell(paragraph("00000001", "A")),
                  cell(nestedTable, paragraph("00000002")),
                ],
              },
            ],
          },
          paragraph("00000003"),
        ],
      },
    },
  } satisfies Document;
  const result = applyDocumentOp(original, {
    type: DOCUMENT_OP_TYPES.MERGE_CELLS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    top: 0,
    bottom: 1,
    left: 0,
    right: 2,
    newBlockIds: Array.from({ length: 1601 }, (_, index) =>
      destinationId(0x800 + index),
    ),
    revision: { id: 100, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
    newIds: { revision: [101, 102, 103, 104, 105, 106, 107] },
  });
  expect(result.isErr()).toBe(true);
  if (result.isOk()) throw new Error("Expected a size refusal.");
  expect(result.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
  );
  expect(result.error.message).toContain("exceeds the saved history limit");
});
