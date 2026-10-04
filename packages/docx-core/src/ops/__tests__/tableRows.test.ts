import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, Table, TableCell, TableRow } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { revisionIdDemand } from "../plan";
import { permitsCellFinalMark } from "../tableTracking";
import { storyParagraphs } from "../blocks";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS, type DocumentOp } from "../types";

const makeParagraph = (
  paraId: string,
  text = "cell",
  fields: Partial<Paragraph> = {},
): Paragraph => ({
  ...fields,
  type: "paragraph",
  paraId,
  content: text.length === 0 ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const makeRow = (paraId: string): TableRow => ({
  type: "tableRow",
  cells: [{ type: "tableCell", content: [makeParagraph(paraId)] }],
});
const tableOf = (...rows: TableRow[]): Table => ({ type: "table", rows });
const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [makeParagraph("00000001", "before"), table, makeParagraph("00000002", "after")],
    },
  },
});
const revision = { id: 100, author: "Reviewer", date: "2026-05-06T07:08:09Z" };
const newIds = { revision: Array.from({ length: 32 }, (_, index) => 101 + index) };
const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return result.value;
};

const expectRefusalWithoutMutation = (document: Document, op: DocumentOp, reason: string): void => {
  const before = structuredClone(document);
  const result = applyDocumentOp(document, op);
  if (result.isOk()) throw new Error(`Expected ${reason} refusal.`);
  expect(result.error.reason).toBe(reason);
  expect(document).toStrictEqual(before);
};

const firstParagraph = (row: TableRow): Paragraph => {
  const paragraph = row.cells.at(0)?.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Fixture row has a first paragraph.");
  return paragraph;
};

const firstCell = (row: TableRow): TableCell => {
  const cell = row.cells.at(0);
  if (cell === undefined) throw new Error("Fixture row has a first cell.");
  return cell;
};

const rowAt = (table: Table, index: number): TableRow => {
  const row = table.rows.at(index);
  if (row === undefined) throw new Error(`Fixture table has row ${index}.`);
  return row;
};

describe("table row review", () => {
  test("resolves a row insertion without changing cell paragraph marks", () => {
    const inserted: TableRow = {
      ...makeRow("00000011"),
      structuralChange: { type: "tableRowInsertion", info: revision },
    };
    const document = documentOf(tableOf(makeRow("00000010"), inserted));
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [100],
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(result.document).toStrictEqual(documentOf(tableOf(makeRow("00000010"))));
  });

  test("insertion at either edge preserves body paragraphs and cell final marks", () => {
    const mark = { kind: "ins" as const, info: { id: 7, author: "Other" } };
    const existing = makeRow("00000010");
    const finalParagraph = firstParagraph(existing);
    const markedRow: TableRow = {
      ...existing,
      cells: [
        {
          ...firstCell(existing),
          content: [{ ...finalParagraph, pPrMark: mark }],
        },
      ],
    };
    const document = documentOf(tableOf(markedRow));
    const [before, , after] = document.package.document.content;
    if (before?.type !== "paragraph" || after?.type !== "paragraph")
      throw new Error("Fixture body paragraphs exist.");

    for (const at of [0, 1]) {
      const inserted = makeRow(at === 0 ? "00000011" : "00000012");
      const result = applied(document, {
        type: DOCUMENT_OP_TYPES.INSERT_ROW,
        story: OP_STORIES.MAIN,
        blockId: "00000010",
        at,
        row: inserted,
      });
      const rows = result.document.package.document.content.find((block) => block.type === "table");
      if (rows?.type !== "table") throw new Error("Fixture table remains present.");
      expect(rowAt(rows, at)).toStrictEqual(inserted);
      expect(before).toBe(result.document.package.document.content[0]);
      expect(after).toBe(result.document.package.document.content[2]);
      expect(firstParagraph(rowAt(rows, at === 0 ? 1 : 0)).pPrMark).toStrictEqual(mark);
    }
  });

  test("tracked insertion in an empty cell records a separate final mark", () => {
    const empty: TableRow = {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [makeParagraph("00000011", "")] }],
    };
    const document = documentOf(tableOf(makeRow("00000010")));
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      at: 1,
      row: empty,
      revision,
      newIds: { revision: [101] },
    });
    expect(result.revisions).toEqual([revision.id, 101]);
    const table = result.document.package.document.content.find((block) => block.type === "table");
    if (table?.type !== "table") throw new Error("Fixture table remains present.");
    expect(rowAt(table, 1)).toStrictEqual({
      ...empty,
      cells: [
        {
          ...firstCell(empty),
          content: [
            { ...firstParagraph(empty), pPrMark: { kind: "ins", info: { ...revision, id: 101 } } },
          ],
        },
      ],
      structuralChange: { type: "tableRowInsertion", info: revision },
    });
  });

  test("deleting a row wraps a foreign insertion and rejecting the row change restores it", () => {
    const foreign = {
      type: "insertion" as const,
      info: { id: 50, author: "Other", date: "2026-01-02T03:04:05Z" },
      content: [{ type: "run" as const, content: [{ type: "text" as const, text: "foreign" }] }],
    };
    const paragraphWithForeign = makeParagraph("00000010", "");
    const sourceRow: TableRow = {
      ...makeRow("00000010"),
      cells: [{ type: "tableCell", content: [{ ...paragraphWithForeign, content: [foreign] }] }],
    };
    const document = documentOf(tableOf(sourceRow, makeRow("00000020")));
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      expected: sourceRow,
      revision,
      newIds,
    });
    const table = result.document.package.document.content.find((block) => block.type === "table");
    if (table?.type !== "table") throw new Error("Fixture table remains present.");
    const trackedParagraph = firstParagraph(rowAt(table, 0));
    expect(trackedParagraph.content).toEqual([
      {
        ...foreign,
        content: [
          {
            type: "deletion",
            info: { ...revision, id: 102 },
            content: foreign.content,
          },
        ],
      },
    ]);
    expect(trackedParagraph.pPrMark).toStrictEqual({
      kind: "del",
      info: { ...revision, id: 101 },
    });
    const rejected = applyDocumentOp(result.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: result.revisions,
      decision: REVISION_DECISIONS.REJECT,
    });
    if (rejected.isErr()) throw rejected.error;
    expect(rejected.value.document).toStrictEqual(document);
  });

  test("tracked deletion of the final row accepts to table removal and rejects exactly", () => {
    const row = makeRow("00000010");
    const document = documentOf(tableOf(row));
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      expected: row,
      revision,
      newIds,
    });
    for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
      const resolved = applied(result.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: result.revisions,
        decision,
      });
      expect(resolved.document).toStrictEqual(
        decision === REVISION_DECISIONS.REJECT
          ? document
          : {
              package: {
                document: {
                  content: [
                    makeParagraph("00000001", "before"),
                    makeParagraph("00000002", "after"),
                  ],
                },
              },
            },
      );
    }
  });

  test("row identities allocate deterministically around existing and repeated ids", () => {
    const document = documentOf(tableOf(makeRow("00000010")));
    document.package.document.content.unshift(
      makeParagraph("00000003", "", {
        pPrMark: { kind: "ins", info: { id: 50, author: "Other" } },
      }),
    );
    const op = {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      at: 1,
      row: makeRow("00000020"),
      revision,
      newIds: { revision: [100, 50, 101, 101, 102] },
    } as const satisfies DocumentOp;
    const result = applied(document, op);
    expect(result.revisions).toEqual([100, 101, 102]);
    expect(applied(structuredClone(document), op).document).toStrictEqual(result.document);
  });

  test.each([
    ["ins", DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT],
    ["del", DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT],
    ["moveFrom", DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE],
    ["moveTo", DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE],
  ] as const)(
    "a pre-existing cell-final %s mark refuses a tracked row operation",
    (kind, reason) => {
      const row: TableRow = {
        ...makeRow("00000010"),
        cells: [
          {
            type: "tableCell",
            content: [
              makeParagraph("00000010", "", {
                pPrMark: { kind, info: { id: 50, author: "Other" } },
              }),
            ],
          },
        ],
      };
      const document = documentOf(tableOf(row));
      expectRefusalWithoutMutation(
        document,
        {
          type: DOCUMENT_OP_TYPES.DELETE_ROW,
          story: OP_STORIES.MAIN,
          blockId: "00000010",
          revision,
          newIds,
        },
        reason,
      );
    },
  );

  test("direct empty row insertion retains an unmarked paragraph and rejects a bare final mark", () => {
    const empty: TableRow = {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [makeParagraph("00000020", "")] }],
    };
    const document = documentOf(tableOf(makeRow("00000010")));
    const op = {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      at: 1,
      row: empty,
    } as const satisfies DocumentOp;
    const result = applied(document, op);
    expect(result.revisions).toEqual([]);
    expect(
      storyParagraphs(result.document.package.document).find(
        ({ paragraph }) => paragraph.paraId === "00000020",
      )?.paragraph,
    ).toStrictEqual(firstParagraph(empty));
    expectRefusalWithoutMutation(
      document,
      {
        ...op,
        row: {
          ...empty,
          cells: [
            {
              ...firstCell(empty),
              content: [{ ...firstParagraph(empty), pPrMark: { kind: "ins", info: revision } }],
            },
          ],
        },
      },
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
    );
  });

  test("cell-final mark ownership follows the innermost row and checks kind and metadata", () => {
    const inner = makeRow("00000031");
    const outer: TableRow = {
      type: "tableRow",
      structuralChange: { type: "tableRowInsertion", info: revision },
      cells: [
        {
          type: "tableCell",
          content: [tableOf(inner), makeParagraph("00000030")],
        },
      ],
    };
    for (const owner of ["outer", "inner"] as const) {
      for (const mismatch of ["none", "author", "date", "initials", "kind", "id"] as const) {
        const row = structuredClone(inner);
        if (owner === "inner") {
          row.structuralChange = { type: "tableRowInsertion", info: revision };
        }
        const mark = {
          kind: mismatch === "kind" ? "del" : "ins",
          info: {
            ...revision,
            id: mismatch === "id" ? 100 : 101,
            author: mismatch === "author" ? "Other" : revision.author,
            date: mismatch === "date" ? "2026-01-01T00:00:00Z" : revision.date,
            ...(mismatch === "initials" ? { initials: "X" } : {}),
          },
        } as const;
        firstParagraph(row).pPrMark = mark;
        const nested = {
          ...outer,
          cells: [
            {
              ...firstCell(outer),
              content: [tableOf(row), makeParagraph("00000030")],
            },
          ],
        };
        const body = documentOf(tableOf(nested)).package.document;
        const location = storyParagraphs(body).find(
          ({ paragraph }) => paragraph.paraId === "00000031",
        );
        if (location === undefined) throw new Error("Fixture nested paragraph exists.");
        expect(permitsCellFinalMark(body, location)).toBe(owner === "inner" && mismatch === "none");
      }
    }
  });

  test("an inverse is stale when another row changes", () => {
    const document = documentOf(tableOf(makeRow("00000010"), makeRow("00000020")));
    const inserted = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      at: 1,
      row: makeRow("00000030"),
    });
    const table = inserted.document.package.document.content.find(
      (block) => block.type === "table",
    );
    if (table?.type !== "table") throw new Error("Fixture table remains present.");
    const changedRows = table.rows.map((row, index) =>
      index === 0 ? { ...row, formatting: { cantSplit: true } } : row,
    );
    const changed = applyDocumentOp(inserted.document, {
      type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      expected: table.rows,
      rows: changedRows,
    });
    if (changed.isErr()) throw changed.error;
    const inverse = inserted.inverse.at(0);
    if (inverse === undefined) throw new Error("Insertion has an inverse.");
    expectRefusalWithoutMutation(
      changed.value.document,
      inverse,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
    );
  });

  test("revision demand distinguishes missing, insufficient and invalid row ids", () => {
    const inserted = makeRow("00000020");
    const secondParagraph = makeParagraph("00000021", "second");
    const twoParagraphRow: TableRow = {
      ...inserted,
      cells: [{ ...firstCell(inserted), content: [firstParagraph(inserted), secondParagraph] }],
    };
    const document = documentOf(tableOf(makeRow("00000010")));
    const op: DocumentOp = {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      at: 1,
      row: twoParagraphRow,
      revision,
    };
    const demand = revisionIdDemand(document, op);
    expect(demand.isOk() ? demand.value : undefined).toBe(3);
    expectRefusalWithoutMutation(document, op, DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS);
    expectRefusalWithoutMutation(
      document,
      { ...op, newIds: { revision: [101] } },
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
    expectRefusalWithoutMutation(
      document,
      { ...op, newIds: { revision: [0x8000_0000, 102, 103] } },
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
    );
  });

  test("a row already carrying a structural revision refuses a new tracked revision", () => {
    const existing: TableRow = {
      ...makeRow("00000010"),
      structuralChange: {
        type: "tableRowInsertion",
        info: { id: 50, author: "Other" },
      },
    };
    const document = documentOf(tableOf(existing, makeRow("00000020")));
    expectRefusalWithoutMutation(
      document,
      {
        type: DOCUMENT_OP_TYPES.DELETE_ROW,
        story: OP_STORIES.MAIN,
        blockId: "00000010",
        revision,
        newIds,
      },
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    );
  });

  test("a paragraph outside a table cannot target a row operation", () => {
    const document = documentOf(tableOf(makeRow("00000010"), makeRow("00000020")));
    expectRefusalWithoutMutation(
      document,
      {
        type: DOCUMENT_OP_TYPES.DELETE_ROW,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
      },
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
  });

  test("a paragraph in a nested table addresses that table and preserves a sibling table", () => {
    const inner = tableOf(makeRow("00000031"));
    const outerRow: TableRow = {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [inner, makeParagraph("00000030")] }],
    };
    const outer = tableOf(outerRow);
    const sibling = tableOf(makeRow("00000040"));
    const before = makeParagraph("00000001", "before");
    const after = makeParagraph("00000002", "after");
    const document: Document = {
      package: { document: { content: [before, outer, sibling, after] } },
    };
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000031",
      at: 1,
      row: makeRow("00000032"),
    });
    const [nextBefore, nextOuter, nextSibling, nextAfter] =
      result.document.package.document.content;
    if (
      nextBefore?.type !== "paragraph" ||
      nextOuter?.type !== "table" ||
      nextSibling?.type !== "table" ||
      nextAfter?.type !== "paragraph"
    ) {
      throw new Error("Fixture content shape is preserved.");
    }
    const nextInner = nextOuter.rows
      .at(0)
      ?.cells.at(0)
      ?.content.find((block) => block.type === "table");
    if (nextInner?.type !== "table") throw new Error("Nested table remains present.");
    expect(nextInner.rows.map((candidate) => firstParagraph(candidate).paraId)).toEqual([
      "00000031",
      "00000032",
    ]);
    expect(nextBefore).toBe(before);
    expect(nextSibling).toBe(sibling);
    expect(nextAfter).toBe(after);
  });

  test.each(["cell", "table property"] as const)(
    "resolving a %s revision clears its selected mark",
    (kind) => {
      const base = makeRow("00000010");
      const markedRow: TableRow =
        kind === "cell"
          ? {
              ...base,
              cells: [
                {
                  ...firstCell(base),
                  structuralChange: {
                    type: "tableCellInsertion",
                    info: { id: 50, author: "Other" },
                  },
                },
              ],
            }
          : base;
      const table = tableOf(markedRow);
      if (kind === "table property") {
        table.formatting = { styleId: "current" };
        table.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: { id: 50, author: "Other" },
            previousFormatting: { styleId: "prior" },
          },
        ];
      }
      const document = documentOf(table);
      const result = applied(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [50],
        decision: REVISION_DECISIONS.ACCEPT,
      });
      const resolvedTable = result.document.package.document.content.find(
        (block) => block.type === "table",
      );
      if (resolvedTable?.type !== "table") throw new Error("Resolved table remains present.");
      if (kind === "cell") {
        expect(resolvedTable.rows.at(0)?.cells.at(0)?.structuralChange).toBeUndefined();
        expect(resolvedTable.rows.at(0)?.cells.at(0)).toBeDefined();
      } else {
        expect(resolvedTable.propertyChanges).toBeUndefined();
        expect(resolvedTable.formatting?.styleId).toBe("current");
      }
    },
  );
});
