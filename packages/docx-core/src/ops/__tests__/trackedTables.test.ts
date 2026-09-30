import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { identityKeysIn } from "../ids";
import { revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DocumentOp,
  type RevisionDecision,
} from "../types";

const makeParagraph = (paraId: string, text = "source"): Paragraph => ({
  type: "paragraph",
  paraId,
  formatting: { alignment: "end", runProperties: { bold: true } },
  textId: "00000070",
  preservedAttributes: [
    {
      namespace: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      name: "rsidR",
      value: "00000080",
    },
  ],
  content:
    text === ""
      ? []
      : [{ type: "run", formatting: { italic: true }, content: [{ type: "text", text }] }],
});
const makeTable = (): Table => ({
  type: "table",
  formatting: { justification: "center" },
  rows: [
    {
      type: "tableRow",
      cells: [
        {
          type: "tableCell",
          content: [makeParagraph("00000010", "first"), makeParagraph("00000011", "last")],
        },
        { type: "tableCell", content: [makeParagraph("00000012", "")] },
      ],
    },
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [makeParagraph("00000020", "second row")] }],
    },
  ],
});
const documentOf = (content: BlockContent[]): Document => ({ package: { document: { content } } });
const revision = { id: 100, author: "Reviewer", date: "2026-05-06T07:08:09Z" };
const newIds = { revision: Array.from({ length: 32 }, (_revisionValue, index) => 101 + index) };
const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return result.value;
};
const resolve = (document: Document, revisionIds: readonly number[], decision: RevisionDecision) =>
  applied(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
const refused = (document: Document, op: DocumentOp, reason: string) => {
  const snapshot = structuredClone(document);
  const result = applyDocumentOp(document, op);
  if (result.isOk()) throw new Error(`Expected ${reason} refusal.`);
  expect(result.error.reason).toBe(reason);
  expect(document).toStrictEqual(snapshot);
};
const firstTable = (document: Document) => {
  const value = document.package.document.content.find((block) => block.type === "table");
  if (value?.type !== "table") throw new Error("Fixture table remains present.");
  return value;
};
const trackedInsertion = (value = makeTable()) =>
  ({
    type: DOCUMENT_OP_TYPES.INSERT_TABLE,
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: "00000002" },
    table: value,
    revision,
    newIds,
  }) as const satisfies DocumentOp;

// These fixtures exercise the fields previously ignored by whole-table operations.
describe("tracked whole tables", () => {
  test("insertion tracks every row, cell content and final cell mark without changing surrounding paragraphs", () => {
    const before = makeParagraph("00000001");
    const final = makeParagraph("00000002", "final");
    const document = documentOf([before, final]);
    const op = trackedInsertion();
    const result = applied(document, op);
    const inserted = firstTable(result.document);
    expect(result.document.package.document.content.at(0)).toBe(before);
    expect(result.document.package.document.content.at(-1)).toBe(final);
    expect(before.pPrMark).toBeUndefined();
    expect(final.pPrMark).toBeUndefined();
    for (const row of inserted.rows) {
      expect(row.structuralChange?.type).toBe("tableRowInsertion");
      for (const cell of row.cells) {
        const last = cell.content.at(-1);
        if (last?.type !== "paragraph") throw new Error("Cells end in paragraphs.");
        expect(last.pPrMark?.kind).toBe("ins");
        for (const block of cell.content) {
          if (block.type !== "paragraph") throw new Error("Fixture cells contain paragraphs.");
          if (block.content.length !== 0) expect(block.content.at(0)?.type).toBe("insertion");
          if (block !== last) expect(block.pPrMark).toBeUndefined();
        }
      }
    }
    const ids = identityKeysIn(inserted);
    expect(ids.length).toBe(8);
    expect(new Set(ids).size).toBe(ids.length);
    expect(new Set(result.revisions)).toEqual(
      new Set(ids.map((key) => Number(key.slice("revision:".length)))),
    );
    expect(
      resolve(result.document, result.revisions, REVISION_DECISIONS.REJECT).document,
    ).toStrictEqual(document);
    const direct = applied(document, {
      type: op.type,
      story: op.story,
      at: op.at,
      table: op.table,
    });
    expect(
      resolve(result.document, result.revisions, REVISION_DECISIONS.ACCEPT).document,
    ).toStrictEqual(direct.document);
  });

  test("C3 insertion after an existing paragraph creates no paragraph mark on its anchor", () => {
    const before = makeParagraph("00000001");
    const final = makeParagraph("00000002", "");
    const document = documentOf([before, final]);
    const result = applied(document, {
      ...trackedInsertion(),
      at: { type: "after", blockId: "00000001" },
    });
    expect(result.document.package.document.content.at(0)).toBe(before);
    expect(before.pPrMark).toBeUndefined();
    expect(
      resolve(result.document, result.revisions, REVISION_DECISIONS.REJECT).document,
    ).toStrictEqual(document);
  });

  test.each([DOCUMENT_OP_TYPES.DELETE_TABLE, DOCUMENT_OP_TYPES.DELETE_ROW] as const)(
    "%s supports tracked deletion of the last row",
    (type) => {
      const value = makeTable();
      const row = value.rows.at(0);
      if (row === undefined) throw new Error("Fixture has a first row.");
      const single = { ...value, rows: [row] };
      const document = documentOf([
        makeParagraph("00000001"),
        single,
        makeParagraph("00000002", ""),
      ]);
      const op = {
        type,
        story: OP_STORIES.MAIN,
        blockId: "00000010",
        revision,
        newIds,
      } as const satisfies DocumentOp;
      const result = applied(document, op);
      const pending = firstTable(result.document);
      expect(pending.rows.at(0)?.structuralChange?.type).toBe("tableRowDeletion");
      for (const cell of pending.rows.at(0)?.cells ?? []) {
        const last = cell.content.at(-1);
        if (last?.type !== "paragraph") throw new Error("Cell ends in a paragraph.");
        expect(last.pPrMark?.kind).toBe("del");
      }
      const direct = applied(document, { type, story: OP_STORIES.MAIN, blockId: "00000010" });
      expect(
        resolve(result.document, result.revisions, REVISION_DECISIONS.ACCEPT).document,
      ).toStrictEqual(direct.document);
      expect(
        resolve(result.document, result.revisions, REVISION_DECISIONS.REJECT).document,
      ).toStrictEqual(document);
    },
  );

  test("revision demand is minimal, distinct and includes empty cell final marks", () => {
    const document = documentOf([makeParagraph("00000002")]);
    const op = trackedInsertion();
    const demand = revisionIdDemand(document, op);
    if (demand.isErr()) throw demand.error;
    expect(demand.value).toBe(7);
    const exact = { ...op, newIds: { revision: newIds.revision.slice(0, demand.value) } };
    const result = applied(document, exact);
    expect(result.revisions.length).toBe(demand.value + 1);
    refused(
      document,
      { ...exact, newIds: { revision: exact.newIds.revision.slice(0, -1) } },
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
    refused(
      document,
      { ...exact, newIds: { revision: exact.newIds.revision.map(() => 101) } },
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
    refused(
      document,
      { ...exact, newIds: { revision: [revision.id, ...exact.newIds.revision.slice(1)] } },
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
  });

  test.each([REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT])(
    "resolving retained rows then final cell marks has exact inverses (%s)",
    (decision) => {
      const document = documentOf([makeParagraph("00000002")]);
      const type =
        decision === REVISION_DECISIONS.ACCEPT
          ? DOCUMENT_OP_TYPES.INSERT_TABLE
          : DOCUMENT_OP_TYPES.DELETE_TABLE;
      const result =
        type === DOCUMENT_OP_TYPES.INSERT_TABLE
          ? applied(document, trackedInsertion())
          : applied(documentOf([makeTable(), makeParagraph("00000002")]), {
              type,
              story: OP_STORIES.MAIN,
              blockId: "00000010",
              revision,
              newIds,
            });
      const rows = firstTable(result.document).rows;
      const rowIds = rows.map((row) => {
        if (row.structuralChange === undefined) throw new Error("Each row is tracked.");
        return row.structuralChange.info.id;
      });
      const retained = resolve(result.document, rowIds, decision);
      expect(
        firstTable(retained.document).rows.every((row) => row.structuralChange === undefined),
      ).toBe(true);
      const cellMarkIds = firstTable(retained.document).rows.flatMap((row) =>
        row.cells.map((cell) => {
          const final = cell.content.at(-1);
          if (final?.type !== "paragraph" || final.pPrMark === undefined)
            throw new Error("Final cell mark survives row resolution.");
          return final.pPrMark.info.id;
        }),
      );
      const cleared = resolve(retained.document, cellMarkIds, decision);
      expect(
        identityKeysIn(cleared.document).filter((key) =>
          cellMarkIds.includes(Number(key.slice("revision:".length))),
        ),
      ).toEqual([]);
    },
  );
});

describe("terminal table insertion carrier", () => {
  test.each(["direct", "tracked"] as const)(
    "%s preserves the authored final paragraph identity and restores exact source",
    (mode) => {
      const final = makeParagraph("00000001", "authored final");
      const document = documentOf([final]);
      const op = {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: "00000001" },
        table: makeTable(),
        terminal: { beforeBlockId: "00000003" },
        ...(mode === "tracked" ? { revision, newIds } : {}),
      } as const satisfies DocumentOp;
      const result = applied(document, op);
      const blocks = result.document.package.document.content;
      expect(blocks.length).toBe(3);
      const preceding = blocks.at(0);
      const trailing = blocks.at(-1);
      if (preceding?.type !== "paragraph" || trailing?.type !== "paragraph")
        throw new Error("Table has both carrier paragraphs.");
      expect(preceding.paraId).toBe("00000003");
      expect(preceding.content).toStrictEqual(final.content);
      expect(preceding.formatting).toStrictEqual(final.formatting);
      expect(trailing).toStrictEqual({ ...final, content: [] });
      expect(trailing.pPrMark).toBeUndefined();
      if (mode !== "tracked") return;
      expect(preceding.pPrMark?.kind).toBe("ins");
      const demand = revisionIdDemand(document, op);
      if (demand.isErr()) throw demand.error;
      expect(demand.value).toBe(8);
      expect(
        resolve(result.document, result.revisions, REVISION_DECISIONS.REJECT).document,
      ).toStrictEqual(document);
      const direct = applied(document, {
        type: op.type,
        story: op.story,
        at: op.at,
        table: op.table,
        terminal: op.terminal,
      });
      expect(
        resolve(result.document, result.revisions, REVISION_DECISIONS.ACCEPT).document,
      ).toStrictEqual(direct.document);
    },
  );

  test("C2 table replacement accepts with only the required final paragraph and rejects original content", () => {
    const original = makeParagraph("00000001", "replace me");
    const document = documentOf([original]);
    const deletion = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: "replace me".length },
    } as const satisfies DocumentOp;
    const removed = applied(document, { ...deletion, revision: { ...revision, id: 90 } });
    const inserted = applied(removed.document, {
      ...trackedInsertion(),
      at: { type: "before", blockId: "00000001" },
    });
    const ids = [...removed.revisions, ...inserted.revisions];
    expect(resolve(inserted.document, ids, REVISION_DECISIONS.REJECT).document).toStrictEqual(
      document,
    );
    const directDeletion = applied(document, deletion);
    const direct = applied(directDeletion.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
      story: OP_STORIES.MAIN,
      at: { type: "before", blockId: "00000001" },
      table: makeTable(),
    });
    const accepted = resolve(inserted.document, ids, REVISION_DECISIONS.ACCEPT);
    expect(accepted.document).toStrictEqual(direct.document);
    expect(accepted.document.package.document.content.length).toBe(2);
    expect(accepted.document.package.document.content.at(-1)).toStrictEqual({
      ...original,
      content: [],
    });
  });

  test("rejecting the preceding break before its inserted table refuses; removing the table first restores the exact source", () => {
    const final = makeParagraph("00000001", "authored final");
    const original = documentOf([final]);
    const inserted = applied(original, {
      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      table: makeTable(),
      terminal: { beforeBlockId: "00000003" },
      revision,
      newIds,
    });
    const preceding = inserted.document.package.document.content.at(0);
    if (preceding?.type !== "paragraph" || preceding.pPrMark === undefined)
      throw new Error("Terminal preceding paragraph has an inserted break.");
    const markId = preceding.pPrMark.info.id;
    refused(
      inserted.document,
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [markId],
        decision: REVISION_DECISIONS.REJECT,
      },
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
    const rowIds = firstTable(inserted.document).rows.map((row) => {
      if (row.structuralChange === undefined) throw new Error("Inserted rows are tracked.");
      return row.structuralChange.info.id;
    });
    const tableRemoved = resolve(inserted.document, rowIds, REVISION_DECISIONS.REJECT);
    expect(tableRemoved.document.package.document.content.length).toBe(2);
    const restored = resolve(tableRemoved.document, [markId], REVISION_DECISIONS.REJECT);
    expect(restored.document).toStrictEqual(original);
  });

  test("terminal carriers reject wrong placement and colliding identities without mutation", () => {
    const document = documentOf([makeParagraph("00000001"), makeParagraph("00000002")]);
    const base = {
      ...trackedInsertion(),
      at: { type: "after", blockId: "00000002" },
      terminal: { beforeBlockId: "00000003" },
    } as const satisfies DocumentOp;
    refused(
      document,
      { ...base, at: { type: "after", blockId: "00000001" } },
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    refused(
      document,
      { ...base, at: { type: "before", blockId: "00000002" } },
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    refused(
      document,
      { ...base, terminal: { beforeBlockId: "00000001" } },
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
    );
    refused(
      document,
      { ...base, terminal: { beforeBlockId: "00000010" } },
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
    );
    refused(
      document,
      { ...base, terminal: { beforeBlockId: "00000000" } },
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
    );
    refused(
      document,
      { ...trackedInsertion(), at: base.at },
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
  });
});

describe("tracked table refusals", () => {
  test.each(["insert", "delete"] as const)(
    "%s refuses unsupported nested/cell structures and existing revisions",
    (mode) => {
      const source = makeTable();
      const row = source.rows.at(0);
      const cell = row?.cells.at(0);
      if (row === undefined || cell === undefined) throw new Error("Fixture has a row and cell.");
      const foreign = { id: 50, author: "Other" };
      const cases = [
        {
          value: {
            ...source,
            rows: [{ ...row, structuralChange: { type: "tableRowInsertion", info: foreign } }],
          },
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
        },
        {
          value: {
            ...source,
            rows: [
              {
                ...row,
                cells: [
                  { ...cell, structuralChange: { type: "tableCellDeletion", info: foreign } },
                ],
              },
            ],
          },
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        },
        {
          value: {
            ...source,
            rows: [
              {
                ...row,
                cells: [
                  {
                    ...cell,
                    content: [
                      {
                        type: "table",
                        rows: [
                          {
                            type: "tableRow",
                            cells: [{ type: "tableCell", content: [makeParagraph("00000030")] }],
                          },
                        ],
                      },
                      makeParagraph("00000010"),
                    ],
                  },
                ],
              },
            ],
          },
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        },
        {
          value: {
            ...source,
            rows: [
              {
                ...row,
                cells: [
                  {
                    ...cell,
                    content: [
                      { ...makeParagraph("00000010"), pPrMark: { kind: "ins", info: foreign } },
                    ],
                  },
                ],
              },
            ],
          },
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
        },
        ...(["moveFrom", "moveTo"] as const).map((type) => ({
          value: {
            ...source,
            rows: [
              {
                ...row,
                cells: [
                  {
                    ...cell,
                    content: [
                      {
                        ...makeParagraph("00000010"),
                        content: [
                          {
                            type,
                            info: foreign,
                            content: [
                              {
                                type: "run" as const,
                                content: [{ type: "text" as const, text: "moved" }],
                              },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        })),
      ] as const satisfies readonly { value: Table; reason: string }[];
      for (const { value, reason } of cases) {
        const document = documentOf(
          mode === "insert" ? [makeParagraph("00000002")] : [value, makeParagraph("00000002")],
        );
        const op =
          mode === "insert"
            ? trackedInsertion(value)
            : ({
                type: DOCUMENT_OP_TYPES.DELETE_TABLE,
                story: OP_STORIES.MAIN,
                blockId: "00000010",
                revision,
                newIds,
              } as const satisfies DocumentOp);
        refused(document, op, reason);
      }
    },
  );

  test("exact review primitives restore cell-ending marks while story-ending mark creation remains refused", () => {
    const value = makeTable();
    const document = documentOf([value, makeParagraph("00000002")]);
    const formatting = makeParagraph("00000011").formatting;
    const review = { formatting, pPrMark: { kind: "ins", info: revision } } as const;
    applied(document, {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story: OP_STORIES.MAIN,
      blockId: "00000011",
      expected: { formatting },
      review,
    });
    refused(
      document,
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
        story: OP_STORIES.MAIN,
        blockId: "00000002",
        expected: { formatting },
        review,
      },
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
    );
  });
});
