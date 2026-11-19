import assert from "node:assert/strict";
import { test } from "node:test";

import {
  type BatchOrderAddress,
  compareBatchOrder,
  orderBatchOperations,
} from "@stll/docx-core/ops";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import {
  pendingJoinTargets,
  sequentialGroups,
} from "../support/batch-replay-order.ts";
import { openReviewer, packDocument } from "../support/documents.ts";
import { restate } from "../support/metamorphic.ts";
import type { Row } from "../support/oracle.ts";

const MODES = ["direct", "tracked-changes"] as const;
const ORDERS = [
  [0, 1, 2],
  [0, 2, 1],
  [1, 0, 2],
  [1, 2, 0],
  [2, 0, 1],
  [2, 1, 0],
] as const;

const view = (reviewer: Awaited<ReturnType<typeof openReviewer>>) =>
  reviewer.getContent().map(({ text, styleId, headingLevel }) => ({
    text,
    styleId,
    headingLevel,
  }));

type MixOptions = {
  mode: (typeof MODES)[number];
  order: (typeof ORDERS)[number];
  mirrored: boolean;
  inserted: string;
};

/** Apply a batch and the contract's right-to-left replay to independent copies. */
const compareMix = async ({ mode, order, mirrored, inserted }: MixOptions) => {
  const fixture = await packDocument(
    fromMarkdown(
      ["Alpha.", "Bravo.", "Charlie.", "Delta.", "Echo."].join("\n\n"),
    ),
  );
  const batch = await openReviewer(fixture);
  const sequential = await openReviewer(fixture);
  const rows = batch.getContent() as Row[];
  const block = (index: number) => {
    const found = rows.at(index);
    assert.ok(found);
    return found;
  };
  const mix = (
    mirrored
      ? [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: block(2).id,
            text: inserted,
          },
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: block(0).id,
            separator: " ",
          },
          {
            id: "replace",
            type: "replaceBlock",
            blockId: block(3).id,
            text: "Replaced delta.",
          },
        ]
      : [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: block(0).id,
            text: inserted,
          },
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: block(2).id,
            separator: " ",
          },
          {
            id: "replace",
            type: "replaceBlock",
            blockId: block(1).id,
            text: "Replaced bravo.",
          },
        ]
  ) satisfies FolioDocumentOperation[];
  const operations = order.flatMap((index) => mix.at(index) ?? []);
  const request = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations,
  } as const;
  const result = batch.applyDocumentOperations(request, {
    undefinedReferences: "refuse",
  });
  assert.deepEqual(result.skipped, []);

  const groups = orderBatchOperations(
    operations.map((operation, inputIndex) => ({ operation, inputIndex })),
    ({ operation, inputIndex }) => {
      const anchor = rows.findIndex(({ id }) => id === operation.blockId);
      const position =
        operation.type === "mergeBlockWithNext" ||
        operation.type === "insertAfterBlock"
          ? anchor + 1
          : anchor;
      return { type: "content" as const, position, rank: 1, inputIndex };
    },
  ).map(({ operation }) => [operation]);
  if (!mirrored && order.join(",") === "0,1,2") {
    assert.deepEqual(
      groups.flat().map(({ id }) => id),
      ["merge", "replace", "insert"],
      "the replay must not hoist the left insertion ahead of right-hand edits",
    );
  }
  for (const group of groups) {
    const replayOperations = group.flatMap(({ id }) =>
      operations.filter((operation) => operation.id === id),
    );
    const replayed = sequential.applyDocumentOperations(
      { ...request, operations: replayOperations },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(replayed.skipped, []);
  }
  if (mode === "tracked-changes") {
    batch.acceptAll();
    sequential.acceptAll();
  }
  assert.deepEqual(view(sequential), view(batch));
};

test("pinned adjacent insertion, merge and replacement batches replay from pre-batch positions", async () => {
  // Fuzz seed 1101252753 (run 128 from base seed 1101252625).
  for (const mode of MODES) {
    await compareMix({
      mode,
      order: ORDERS[0],
      mirrored: false,
      inserted: "Added 👍🏽 clause.",
    });
  }
});

test("overlapping insertions and merges replay their applied set in every input order", async () => {
  for (const order of ORDERS) {
    const fixture = await packDocument(
      fromMarkdown("Alpha.\n\nBravo.\n\nCharlie."),
    );
    const batch = await openReviewer(fixture);
    const sequential = await openReviewer(fixture);
    const rows = batch.getContent() as Row[];
    const anchor = rows.at(0);
    assert.ok(anchor);
    const operations: FolioDocumentOperation[] = [
      {
        id: "insert-a",
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Inserted A.",
      },
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: anchor.id,
        separator: " ",
      },
      {
        id: "insert-b",
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Inserted B.",
      },
    ];
    const ordered = order.flatMap((index) => operations.at(index) ?? []);
    const request = {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct" as const,
      operations: ordered,
    };
    const result = batch.applyDocumentOperations(request, {
      undefinedReferences: "refuse",
    });
    const appliedIds = new Set(result.applied.map(({ id }) => id));
    const applied = ordered.filter(({ id }) => appliedIds.has(id));
    // Same-anchor insertion and merge claims overlap. Replay only the
    // operations the real batch accepted, accounting for every request.
    assert.deepEqual(
      [...result.applied, ...result.skipped].map(({ id }) => id).toSorted(),
      ordered.map(({ id }) => id).toSorted(),
    );
    assert.ok(
      result.skipped.every(({ reason }) => reason === "overlappingOperation"),
    );
    const groups = sequentialGroups({
      applied,
      preRows: rows,
      mode: "direct",
    });
    assert.deepEqual(
      groups
        .flat()
        .map(({ id }) => id)
        .toSorted(),
      [...appliedIds].toSorted(),
    );
    if (order === ORDERS[0]) {
      assert.deepEqual(
        result.skipped.map(({ id }) => id),
        ["merge"],
      );
      assert.deepEqual(
        groups.map((group) => group.map(({ id }) => id)),
        [["insert-a", "insert-b"]],
      );
    }
    for (const group of groups) {
      const replayed = sequential.applyDocumentOperations(
        { ...request, operations: group as FolioDocumentOperation[] },
        { undefinedReferences: "refuse" },
      );
      assert.deepEqual(replayed.skipped, []);
    }
    assert.deepEqual(view(sequential), view(batch));
  }
});

test("restate refuses replaceBlock forms whose formatting semantics cannot become a range", () => {
  const preRows = [{ id: "block", text: "Original." }] as Row[];
  const rows = [{ id: "block", text: "Prefix Original." }] as Row[];
  const base = {
    id: "replace",
    type: "replaceBlock",
    blockId: "block",
    text: "Replacement.",
  };
  const textHashOf = () => undefined;

  assert.equal(
    restate({
      operation: { ...base, styleId: "Heading1" },
      preRows,
      rows,
      textHashOf,
    }),
    null,
  );
  assert.equal(
    restate({
      operation: { ...base, preserveFormatting: false },
      preRows,
      rows,
      textHashOf,
    }),
    null,
  );
  assert.equal(
    restate({ operation: base, preRows, rows, textHashOf })?.type,
    "replaceRange",
  );
});

test("the batch comparator defines one total order", () => {
  const addresses: BatchOrderAddress[] = [];
  for (let inputIndex = 0; inputIndex < 4; inputIndex++) {
    for (let position = 0; position < 3; position++) {
      addresses.push({ type: "annotation", position, inputIndex });
      for (let rank = 0; rank < 3; rank++) {
        addresses.push({ type: "content", position, rank, inputIndex });
      }
      addresses.push({
        type: "cell",
        tablePosition: position,
        bottom: 1,
        right: 2,
        inputIndex,
      });
      addresses.push({
        type: "column",
        tablePosition: position,
        columnIndex: 1,
        insertion: inputIndex % 2 === 0,
        inputIndex,
      });
    }
  }
  const sign = (value: number) => Math.sign(value);
  for (const left of addresses) {
    for (const right of addresses) {
      const comparison = compareBatchOrder(left, right);
      const reverse = sign(compareBatchOrder(right, left));
      assert.ok(
        sign(comparison) === -reverse || (comparison === 0 && reverse === 0),
      );
      if (comparison === 0) assert.deepEqual(left, right);
      for (const last of addresses) {
        if (comparison <= 0 && compareBatchOrder(right, last) <= 0) {
          assert.ok(compareBatchOrder(left, last) <= 0);
        }
      }
    }
  }
});

test("the batch comparator preserves phases within a table and visits nested tables first", () => {
  const sameTable: BatchOrderAddress[] = [
    {
      type: "column",
      tablePosition: 10,
      columnIndex: 0,
      insertion: false,
      inputIndex: 3,
    },
    { type: "cell", tablePosition: 10, bottom: 1, right: 1, inputIndex: 2 },
    { type: "content", position: 10, rank: 1, inputIndex: 1 },
    { type: "annotation", position: 10, inputIndex: 0 },
  ];
  assert.deepEqual(
    orderBatchOperations(sameTable, (address) => address).map(
      ({ type }) => type,
    ),
    ["annotation", "content", "cell", "column"],
  );

  const outerCell: BatchOrderAddress = {
    type: "cell",
    tablePosition: 10,
    bottom: 2,
    right: 2,
    inputIndex: 0,
  };
  const nestedColumn: BatchOrderAddress = {
    type: "column",
    tablePosition: 20,
    columnIndex: 0,
    insertion: false,
    inputIndex: 1,
  };
  assert.deepEqual(
    orderBatchOperations([outerCell, nestedColumn], (address) => address),
    [nestedColumn, outerCell],
  );
});

test("a table insertion replays before the deletion of its anchor through every three-operation permutation", async () => {
  const outcomes: ReturnType<typeof view>[] = [];
  const split = async () => {
    const fixture = await packDocument(fromMarkdown("Alpha bravo."));
    const reviewer = await openReviewer(fixture);
    const anchor = (reviewer.getContent() as Row[]).at(0);
    assert.ok(anchor);
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          { id: "split", type: "splitBlock", blockId: anchor.id, offset: 5 },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(result.skipped, []);
    return reviewer;
  };
  for (const order of ORDERS) {
    const reviewer = await split();
    const sequential = await split();
    const splitAnchor = (reviewer.getContent() as Row[]).at(0);
    assert.ok(splitAnchor);
    const operations = [
      {
        id: "table",
        type: "insertTable" as const,
        blockId: splitAnchor.id,
        rows: [["Term", "Value"]],
      },
      {
        id: "before",
        type: "insertBeforeBlock" as const,
        blockId: splitAnchor.id,
        text: "Before.",
      },
      { id: "delete", type: "deleteBlock" as const, blockId: splitAnchor.id },
    ];
    const replay = sequentialGroups({
      applied: order.flatMap((index) => operations.at(index) ?? []),
      preRows: reviewer.getContent() as Row[],
      mode: "tracked-changes",
    });
    const orderedReplay = replay.flat().map(({ id }) => id);
    assert.ok(orderedReplay.indexOf("table") < orderedReplay.indexOf("delete"));
    for (const group of replay) {
      const replayed = sequential.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: group as FolioDocumentOperation[],
        },
        { undefinedReferences: "refuse" },
      );
      assert.deepEqual(replayed.skipped, []);
    }
    sequential.acceptAll();
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: order.flatMap((index) => operations.at(index) ?? []),
      },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(result.skipped, []);
    reviewer.acceptAll();
    assert.deepEqual(view(sequential), view(reviewer));
    outcomes.push(view(reviewer));
  }
  for (const outcome of outcomes.slice(1))
    assert.deepEqual(outcome, outcomes.at(0));
});

test("generated adjacent insertion, merge and replacement mixes equal contract replay", async () => {
  const insertedTexts = [
    "Added clause.",
    "Added 👍🏽 clause.",
    "Cláusula añadida.",
    "סעיף נוסף.",
  ];
  for (const [index, order] of ORDERS.entries()) {
    await compareMix({
      mode: MODES[index % MODES.length] ?? "direct",
      order,
      mirrored: index % 2 === 1,
      inserted: insertedTexts[index % insertedTexts.length] ?? "Added clause.",
    });
  }
});

test("a block addressed and taken out in one batch replays the address first, in either input order", async () => {
  const fixture = await packDocument(
    fromMarkdown("Alpha.\n\nBravo.\n\nCharlie."),
  );
  const rows = (await openReviewer(fixture)).getContent() as Row[];
  const [alpha, bravo] = rows;
  assert.ok(alpha && bravo);
  const cases: [string, FolioDocumentOperation, FolioDocumentOperation][] = [
    [
      "deleted",
      {
        id: "address",
        type: "insertTable",
        blockId: bravo.id,
        rows: [["Term", "Value"]],
      },
      { id: "remove", type: "deleteBlock", blockId: bravo.id },
    ],
    [
      "merged away",
      {
        id: "address",
        type: "insertBeforeBlock",
        blockId: bravo.id,
        text: "Before.",
      },
      {
        id: "remove",
        type: "mergeBlockWithNext",
        blockId: alpha.id,
        separator: " ",
      },
    ],
  ];
  for (const [name, address, remove] of cases) {
    for (const operations of [
      [address, remove],
      [remove, address],
    ]) {
      for (const mode of MODES) {
        const replay = sequentialGroups({
          applied: operations,
          preRows: rows,
          mode,
        });
        const order = replay.flat().map(({ id }) => id);
        assert.ok(
          order.indexOf("address") < order.indexOf("remove"),
          `${name}, ${mode}: ${order.join(", ")}`,
        );
        const sequential = await openReviewer(fixture);
        for (const group of replay) {
          const replayed = sequential.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: group as FolioDocumentOperation[],
            },
            { undefinedReferences: "refuse" },
          );
          assert.deepEqual(replayed.skipped, [], `${name}, ${mode}`);
        }
      }
    }
  }
});

/** Apply `operations` as one batch and in the replay's order, to independent copies. */
const compareWithReplay = async (
  fixture: Uint8Array,
  mode: "direct" | "tracked-changes" | "suggested",
  operations: readonly FolioDocumentOperation[],
) => {
  const batched = await openReviewer(fixture);
  const preRows = batched.getContent() as Row[];
  const joinSources = new Set(
    batched
      .getChanges()
      .flatMap(({ type, blockId }) =>
        type === "paragraphMarkDeleted" && blockId !== null ? [blockId] : [],
      ),
  );
  const batch = batched.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [...operations],
    },
    { undefinedReferences: "refuse" },
  );
  assert.deepEqual(batch.skipped, [], `${mode}: batch`);
  const sequential = await openReviewer(fixture);
  for (const group of sequentialGroups({
    applied: operations,
    preRows,
    mode,
    joinsInto:
      mode === "suggested"
        ? new Map()
        : pendingJoinTargets(preRows, joinSources),
  })) {
    const replayed = sequential.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: group as FolioDocumentOperation[],
      },
      { undefinedReferences: "refuse" },
    );
    assert.deepEqual(replayed.skipped, [], `${mode}: one at a time`);
  }
  assert.deepEqual(view(sequential), view(batched), mode);
};

test("a row deletion replays after an insertion anchored in that row", async () => {
  const fixture = await packDocument(
    fromMarkdown("Alpha.\n\n| Term | Value |\n| --- | --- |\n| A | B |"),
  );
  const rows = (await openReviewer(fixture)).getContent() as Row[];
  const cell = rows.find(({ table }) => table !== undefined);
  assert.ok(cell);
  const remove: FolioDocumentOperation = {
    id: "remove",
    type: "deleteTableRow",
    blockId: cell.id,
  };
  const insert: FolioDocumentOperation = {
    id: "insert",
    type: "insertBeforeBlock",
    blockId: cell.id,
    text: "Before.",
  };
  for (const mode of [...MODES, "suggested"] as const) {
    await compareWithReplay(fixture, mode, [remove, insert]);
    await compareWithReplay(fixture, mode, [insert, remove]);
  }
});

test("a sole paragraph's deletion replays first, as the batch runs it", async () => {
  const fixture = await packDocument(fromMarkdown("Only."));
  const rows = (await openReviewer(fixture)).getContent() as Row[];
  const only = rows.at(0);
  assert.ok(only);
  const remove: FolioDocumentOperation = {
    id: "remove",
    type: "deleteBlock",
    blockId: only.id,
  };
  const insert: FolioDocumentOperation = {
    id: "insert",
    type: "insertBeforeBlock",
    blockId: only.id,
    text: "Before.",
  };
  for (const mode of MODES) {
    await compareWithReplay(fixture, mode, [remove, insert]);
    await compareWithReplay(fixture, mode, [insert, remove]);
  }
});

test("a block after a pending join replays ahead of a split of the paragraph it joins", async () => {
  const source = await openReviewer(
    await packDocument(
      fromMarkdown("Alpha.\n\nClause.\n\nSigned in two copies.\n\nEnd."),
    ),
  );
  const [, clause, survivor] = source.getContent();
  assert.ok(clause && survivor);
  source.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: clause.id,
          separator: " ",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const fixture = new Uint8Array(await source.toBuffer());
  const rows = (await openReviewer(fixture)).getContent() as Row[];
  assert.deepEqual(
    pendingJoinTargets(rows, new Set([clause.id])),
    new Map([[clause.id, { survivor: survivor.id, keepsWords: true }]]),
  );
  const insert: FolioDocumentOperation = {
    id: "insert",
    type: "insertTable",
    blockId: clause.id,
    rows: [["Term", "Value"]],
  };
  const split: FolioDocumentOperation = {
    id: "split",
    type: "splitBlock",
    blockId: survivor.id,
    offset: 14,
  };
  for (const mode of MODES) {
    await compareWithReplay(fixture, mode, [insert, split]);
    await compareWithReplay(fixture, mode, [split, insert]);
  }
});

test("a block after a paragraph joining an emptied final one replays ahead of its new words", async () => {
  const source = await openReviewer(
    await packDocument(fromMarkdown("Alpha.\n\nThe Buyer pays.\n\nSigned.")),
  );
  const [, buyer, final] = source.getContent();
  assert.ok(buyer && final);
  // Deleting the final paragraph tracked deletes its words and the break
  // before it: the paragraph before it joins it, which has no words left.
  source.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: final.id }],
    },
    { undefinedReferences: "refuse" },
  );
  const fixture = new Uint8Array(await source.toBuffer());
  const reopened = await openReviewer(fixture);
  const rows = reopened.getContent() as Row[];
  const sources = new Set(
    reopened
      .getChanges()
      .flatMap(({ type, blockId }) =>
        type === "paragraphMarkDeleted" && blockId !== null ? [blockId] : [],
      ),
  );
  assert.deepEqual(
    pendingJoinTargets(rows, sources),
    new Map([[buyer.id, { survivor: final.id, keepsWords: false }]]),
  );
  const insert: FolioDocumentOperation = {
    id: "insert",
    type: "insertAfterBlock",
    blockId: buyer.id,
    text: "Inserted.",
  };
  const replace: FolioDocumentOperation = {
    id: "replace",
    type: "replaceBlock",
    blockId: final.id,
    text: "Replaced.",
  };
  await compareWithReplay(fixture, "tracked-changes", [insert, replace]);
  await compareWithReplay(fixture, "tracked-changes", [replace, insert]);
  // Put before the emptied paragraph, a block lands in the same gap.
  const before: FolioDocumentOperation = {
    id: "before",
    type: "insertBeforeBlock",
    blockId: final.id,
    text: "Before.",
  };
  await compareWithReplay(fixture, "tracked-changes", [before, replace]);
  await compareWithReplay(fixture, "tracked-changes", [replace, before]);
  await compareWithReplay(fixture, "tracked-changes", [
    insert,
    before,
    replace,
  ]);
});
