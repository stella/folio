import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";

import { openReviewer, packDocument } from "../support/documents.ts";
import {
  capture,
  compareComments,
  expectOperation,
  modelOf,
  rowsOf,
} from "../support/oracle.ts";
import { sourceOperationFootprint } from "../support/table-oracle.ts";

const p = (id: string) =>
  ({
    type: "paragraph",
    paraId: id,
    content: [{ type: "run", content: [{ type: "text", text: id }] }],
  }) as const;
const nested = {
  type: "table",
  columnWidths: [1000, 1000],
  rows: [
    {
      type: "tableRow",
      cells: [
        { type: "tableCell", content: [p("10000002"), p("10000003")] },
        { type: "tableCell", content: [p("10000004"), p("10000005")] },
      ],
    },
  ],
} as const;
const input = () =>
  packDocument({
    package: {
      document: {
        content: [
          {
            type: "table",
            columnWidths: [2000, 2000],
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    formatting: { gridSpan: 2 },
                    content: [p("10000001"), nested, p("10000006")],
                  },
                ],
              },
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [p("10000007"), p("10000008")],
                  },
                  {
                    type: "tableCell",
                    content: [p("10000009"), p("1000000A")],
                  },
                ],
              },
            ],
          },
          {
            type: "table",
            columnWidths: [2000],
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [p("1000000B"), p("1000000C")],
                  },
                ],
              },
            ],
          },
          p("1000000D"),
        ],
      },
    },
  });

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Operation = Parameters<
  Reviewer["applyDocumentOperations"]
>[0]["operations"][number];
const apply = (reviewer: Reviewer, operations: readonly Operation[]) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(
    result.applied.length,
    operations.length,
    JSON.stringify(result),
  );
  assert.deepEqual(result.skipped, []);
};

test("source cell footprints preserve nested ancestry and merged-span intersections", async () => {
  const reviewer = await openReviewer(await input());
  const rows = rowsOf(reviewer);
  const wide = rows.find(({ id }) => id === "10000001");
  assert.equal(wide?.table?.columnSpan, 2);
  const cases = [
    {
      operation: { type: "splitTableCell", blockId: "10000001" },
      ids: [
        "10000001",
        "10000002",
        "10000003",
        "10000004",
        "10000005",
        "10000006",
      ],
    },
    {
      operation: { type: "deleteTableColumn", blockId: "10000009" },
      ids: [
        "10000001",
        "10000002",
        "10000003",
        "10000004",
        "10000005",
        "10000006",
        "10000009",
        "1000000A",
      ],
    },
    {
      operation: { type: "deleteTableRow", blockId: "10000001" },
      ids: [
        "10000001",
        "10000002",
        "10000003",
        "10000004",
        "10000005",
        "10000006",
      ],
    },
    {
      operation: {
        type: "mergeTableCells",
        blockId: "10000007",
        endBlockId: "10000009",
      },
      ids: ["10000007", "10000008", "10000009", "1000000A"],
    },
    {
      operation: {
        type: "mergeTableCells",
        blockId: "10000002",
        endBlockId: "10000004",
      },
      ids: ["10000002", "10000003", "10000004", "10000005"],
    },
  ];
  for (const { operation, ids } of cases) {
    const footprint = sourceOperationFootprint({ rows, operation });
    assert.equal(footprint.type, "known");
    if (footprint.type !== "known") assert.fail(footprint.reason);
    assert.deepEqual([...footprint.ids].sort(), ids.sort());
  }
  const missing = sourceOperationFootprint({
    rows: rows.map((row) => Object.assign({}, row)),
    operation: { type: "splitTableCell", blockId: "10000001" },
  });
  assert.equal(
    missing.type,
    "unsupported",
    "Copied/manual rows must not invent source ancestry",
  );
  const missingDescendant = sourceOperationFootprint({
    rows: rows.map((row) =>
      row.id === "10000002" ? Object.assign({}, row) : row,
    ),
    operation: { type: "splitTableCell", blockId: "10000001" },
    requiredParagraphIds: ["10000002"],
  });
  assert.equal(
    missingDescendant.type,
    "unsupported",
    "An unlocated nested join cannot silently retain an exact destination proof",
  );
});

test("an applied parent-cell split expires its nested join but keeps sibling cells and tables exact", async () => {
  const reviewer = await openReviewer(await input());
  apply(reviewer, [
    {
      id: "nested-join",
      type: "mergeBlockWithNext",
      blockId: "10000002",
      separator: " ",
    },
    {
      id: "sibling-join",
      type: "mergeBlockWithNext",
      blockId: "10000007",
      separator: " ",
    },
    {
      id: "other-table-join",
      type: "mergeBlockWithNext",
      blockId: "1000000B",
      separator: " ",
    },
  ]);
  const pre = await capture(reviewer, "direct");
  const operation = {
    id: "split-wide",
    type: "splitTableCell",
    blockId: "10000001",
  } as const;
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(result.applied.length, 1, JSON.stringify(result));
  assert.deepEqual(result.skipped, []);
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  model.mode = pre.mode;
  expectOperation(model, operation);
  assert.equal(model.unmodelled.length, 1);
  assert.deepEqual(
    [...model.pendingJoins.values()]
      .map(({ destination, destinationPrediction }) => [
        destination.id,
        destinationPrediction.type,
      ])
      .sort(),
    [
      ["10000003", "unknown"],
      ["10000008", "exact"],
      ["1000000C", "exact"],
    ],
  );
  assert.equal(
    rowsOf(reviewer).find(({ id }) => id === "10000001")?.table?.columnSpan,
    1,
    "The real producer split the wide parent cell",
  );
});

test("a null cell merge expires source groups beyond its addressed cell and preserves an unrelated table comment", async () => {
  const reviewer = await openReviewer(await input());
  apply(reviewer, [
    {
      id: "left-join",
      type: "mergeBlockWithNext",
      blockId: "10000007",
      separator: " ",
    },
    {
      id: "right-join",
      type: "mergeBlockWithNext",
      blockId: "10000009",
      separator: " ",
    },
    {
      id: "other-join",
      type: "mergeBlockWithNext",
      blockId: "1000000B",
      separator: " ",
    },
  ]);
  apply(reviewer, [
    {
      id: "keep-other",
      type: "commentOnBlock",
      blockId: "1000000C",
      quote: "1000000C",
      comment: { text: "Keep the other table owner." },
    },
  ]);
  const pre = await capture(reviewer, "direct");
  const operation = {
    id: "merge-cells",
    type: "mergeTableCells",
    blockId: "10000007",
    endBlockId: "10000009",
  } as const;
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(result.applied.length, 1, JSON.stringify(result));
  assert.deepEqual(result.skipped, []);
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  model.mode = pre.mode;
  expectOperation(model, operation);
  assert.deepEqual(
    [...model.pendingJoins.values()]
      .map(({ destination, destinationPrediction }) => [
        destination.id,
        destinationPrediction.type,
      ])
      .sort(),
    [
      ["10000008", "unknown"],
      ["1000000A", "unknown"],
      ["1000000C", "exact"],
    ],
  );
  const after = await capture(reviewer, "direct");
  assert.deepEqual(
    compareComments(
      model,
      pre.comments,
      after.comments,
      after.rows,
      pre.liveComments,
      "direct",
    ),
    [],
  );
  const comment = after.comments.at(0);
  assert.ok(comment);
  assert.equal(
    compareComments(
      model,
      pre.comments,
      [{ ...comment, anchor: "drift" }],
      after.rows,
      pre.liveComments,
      "direct",
    ).length,
    1,
  );
  assert.equal(
    compareComments(
      model,
      pre.comments,
      [],
      after.rows,
      pre.liveComments,
      "direct",
    ).length,
    1,
  );
});

test("a paragraph join footprint includes the next source component even outside the addressed join group", async () => {
  const reviewer = await openReviewer(
    await packDocument(fromMarkdown("Outside.\n\nSource.\n\nDestination.")),
  );
  const [outside, source, destination] = rowsOf(reviewer);
  assert.ok(outside && source && destination);
  apply(reviewer, [
    {
      id: "pending-join",
      type: "mergeBlockWithNext",
      blockId: source.id,
      separator: " ",
    },
  ]);
  const pre = await capture(reviewer, "tracked-changes");
  const footprint = sourceOperationFootprint({
    rows: pre.liveRows,
    operation: { type: "mergeBlockWithNext", blockId: outside.id },
  });
  assert.equal(footprint.type, "known");
  if (footprint.type !== "known") assert.fail(footprint.reason);
  assert.deepEqual([...footprint.ids], [outside.id, source.id]);
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  const ownership = [...model.pendingJoins.values()].at(0);
  assert.ok(ownership);
  assert.ok(ownership.sources.some(({ id }) => footprint.ids.has(id)));
  assert.equal(
    ownership.sources.some(({ id }) => id === outside.id),
    false,
  );
});

const verticalInput = () =>
  packDocument({
    package: {
      document: {
        content: [
          {
            type: "table",
            columnWidths: [2000, 2000],
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    formatting: { vMerge: "restart" },
                    content: [p("20000001"), p("20000002")],
                  },
                  { type: "tableCell", content: [p("20000003")] },
                ],
              },
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    formatting: { vMerge: "continue" },
                    content: [
                      { type: "paragraph", paraId: "20000004", content: [] },
                    ],
                  },
                  { type: "tableCell", content: [p("20000005")] },
                ],
              },
            ],
          },
          p("20000006"),
        ],
      },
    },
  });

test("a row footprint includes a source cell whose vertical span starts above the removed row", async () => {
  const reviewer = await openReviewer(await verticalInput());
  const rows = rowsOf(reviewer);
  assert.equal(rows.find(({ id }) => id === "20000001")?.table?.rowSpan, 2);
  const footprint = sourceOperationFootprint({
    rows,
    operation: { type: "deleteTableRow", blockId: "20000005" },
  });
  assert.equal(footprint.type, "known");
  if (footprint.type !== "known") assert.fail(footprint.reason);
  assert.ok(footprint.ids.has("20000001") && footprint.ids.has("20000002"));
  assert.ok(footprint.ids.has("20000005"));
  assert.equal(footprint.ids.has("20000003"), false);
  assert.equal(footprint.ids.has("20000006"), false);
});

for (const spec of [
  {
    type: "insertTableColumn",
    input,
    source: "10000002",
    destination: "10000003",
    anchor: "10000007",
    position: "after",
    spanOwner: "10000001",
    span: "columnSpan",
  },
  {
    type: "insertTableRow",
    input: verticalInput,
    source: "20000001",
    destination: "20000002",
    anchor: "20000005",
    position: "before",
    spanOwner: "20000001",
    span: "rowSpan",
  },
] as const)
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`${mode} ${spec.type} across a spanning cell preserves its known joined owner and quote`, async () => {
      const reviewer = await openReviewer(await spec.input());
      apply(reviewer, [
        {
          id: "join",
          type: "mergeBlockWithNext",
          blockId: spec.source,
          separator: " ",
        },
      ]);
      apply(reviewer, [
        {
          id: "comment",
          type: "commentOnBlock",
          blockId: spec.destination,
          quote: spec.destination,
          comment: { text: "Keep the spanning cell comment." },
        },
      ]);
      const pre = await capture(reviewer, mode);
      const operation = {
        id: "insert-through-span",
        type: spec.type,
        blockId: spec.anchor,
        position: spec.position,
      } as const;
      const footprint = sourceOperationFootprint({
        rows: pre.liveRows,
        operation,
      });
      assert.equal(footprint.type, "known");
      if (footprint.type !== "known") assert.fail(footprint.reason);
      assert.equal(footprint.ids.size, 0);
      const before = reviewer.toDocument();
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [operation],
        },
        { undefinedReferences: "refuse" },
      );
      if (mode === "tracked-changes" && spec.type === "insertTableColumn") {
        assert.deepEqual(result.applied, []);
        assert.deepEqual(
          result.skipped.map(({ id, reason }) => ({ id, reason })),
          [{ id: operation.id, reason: "unsupportedBlock" }],
        );
        assert.deepEqual(reviewer.toDocument(), before);
        return;
      }
      assert.equal(result.applied.length, 1, JSON.stringify(result));
      assert.deepEqual(result.skipped, []);
      const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
      model.mode = pre.mode;
      expectOperation(model, operation);
      assert.ok(
        [...model.pendingJoins.values()].every(
          ({ destinationPrediction }) => destinationPrediction.type === "exact",
        ),
      );
      assert.equal(
        rowsOf(reviewer).find(({ id }) => id === spec.spanOwner)?.table?.[
          spec.span
        ],
        3,
        "The real producer expands the crossing span",
      );
      const after = await capture(reviewer, mode);
      assert.deepEqual(
        compareComments(
          model,
          pre.comments,
          after.comments,
          after.rows,
          pre.liveComments,
          mode,
        ),
        [],
      );
      const comment = after.comments.at(0);
      assert.ok(comment);
      assert.equal(comment.blockId, spec.destination);
      assert.equal(comment.anchor, spec.destination);
      assert.equal(
        compareComments(
          model,
          pre.comments,
          [{ ...comment, anchor: "drift" }],
          after.rows,
          pre.liveComments,
          mode,
        ).length,
        1,
      );
    });

test("outside paragraph run-in across a table never acquires an exact same-container text proof", async () => {
  const reviewer = await openReviewer(
    await packDocument({
      package: {
        document: {
          content: [
            {
              ...p("30000001"),
              content: [
                { type: "run", content: [{ type: "text", text: "Outside" }] },
              ],
              pPrMark: { kind: "del", info: { id: 1, author: "Reviewer" } },
            },
            {
              type: "table",
              columnWidths: [2000, 2000],
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    {
                      type: "tableCell",
                      content: [
                        {
                          ...p("30000002"),
                          content: [
                            {
                              type: "run",
                              content: [{ type: "text", text: "First" }],
                            },
                          ],
                        },
                      ],
                    },
                    {
                      type: "tableCell",
                      content: [
                        {
                          ...p("30000003"),
                          content: [
                            {
                              type: "run",
                              content: [{ type: "text", text: "Second" }],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
            p("30000004"),
          ],
        },
      },
    }),
  );
  const outside = rowsOf(reviewer).find(({ id }) => id === "30000001");
  assert.ok(outside);
  const source = await capture(reviewer, "direct");
  const sourceModel = modelOf(source.rows, source.liveRows, source.joinSources);
  const ownership = [...sourceModel.pendingJoins.values()].find(({ sources }) =>
    sources.some(({ id }) => id === outside.id),
  );
  assert.ok(ownership);
  assert.equal(ownership.sources.at(0)?.table, undefined);
  assert.ok(ownership.destination.table);
  const commentOperation = {
    id: "comment-run-in",
    type: "commentOnBlock",
    blockId: outside.id,
    quote: "Outside",
    comment: { text: "Keep the outside comment." },
  } as const;
  apply(reviewer, [commentOperation]);
  expectOperation(sourceModel, commentOperation);
  assert.equal(sourceModel.unmodelled.length, 1);
  assert.match(
    sourceModel.unmodelled.at(0) ?? "",
    /no exact live-to-accepted text projection/u,
  );

  const pre = await capture(reviewer, "direct");
  const otherCell = pre.liveRows.find(
    ({ text, table }) => text === "Second" && table !== undefined,
  );
  assert.ok(otherCell);
  const operation = {
    id: "before-outer",
    type: "insertTable",
    blockId: otherCell.id,
    position: "before",
    rows: [["New"]],
  } as const;
  const before = reviewer.toDocument();
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  assert.deepEqual(result.applied, []);
  assert.deepEqual(
    result.skipped.map(({ id, reason }) => ({ id, reason })),
    [{ id: operation.id, reason: "pendingParagraphMarkDeletion" }],
  );
  assert.deepEqual(reviewer.toDocument(), before);
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  assert.equal(
    compareComments(
      model,
      pre.comments,
      [],
      pre.rows,
      pre.liveComments,
      "direct",
    ).length,
    1,
    "An unproven run-in mapping does not excuse known comment loss",
  );
});
