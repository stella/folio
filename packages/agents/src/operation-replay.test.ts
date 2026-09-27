/**
 * An operation id names one operation for the whole document session, so a
 * `suggest_changes` call resent after a lost response changes nothing the
 * second time.
 *
 * Every operation type is applied, then replayed with the same id: the replay
 * reports the operation under `replayed` with its original receipt, and the
 * document, its tracked changes and its comments are what the first call left.
 * The builders are keyed by operation type, so a type added to the contract
 * without a replay case does not compile.
 */

import { describe, expect, test } from "bun:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createDocx,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  FolioDocxReviewer,
  getFolioDocumentOperationReceipts,
  isFolioDocumentOperationModeSupported,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "./bridge";
import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCallUntyped } from "./execute";
import type { FolioAgentApplyOperationsSummary } from "./types";

const MARKDOWN = [
  "# Agreement",
  "The supplier delivers the goods on time.",
  "The buyer pays within thirty days.",
  "Either party may terminate on notice.",
  "| Item | Price |\n|---|---|\n| Goods | 100 |\n|  | 200 |",
  "Closing words.",
].join("\n\n");

type Blocks = {
  heading: string;
  first: string;
  second: string;
  third: string;
  cell: (row: number, column: number) => string;
  /** A range over the first word of `first`, as `find_text` returns it. */
  range: unknown;
};

type ReplayCase = {
  /**
   * An operation of another id applied directly first, when the replayed one
   * needs it: a tracked split takes a committed merge, not a pending one.
   */
  setup?: (blocks: Blocks) => Record<string, unknown>;
  operation: (blocks: Blocks) => Record<string, unknown>;
};

const REPLAY_CASES = {
  replaceInBlock: {
    operation: ({ first }) => ({ blockId: first, find: "supplier", replace: "vendor" }),
  },
  replaceRange: { operation: ({ range }) => ({ range, replace: "A" }) },
  commentOnRange: { operation: ({ range }) => ({ range, comment: "Which article?" }) },
  formatRange: { operation: ({ range }) => ({ range, formatting: { bold: true } }) },
  insertAfterBlock: { operation: ({ first }) => ({ blockId: first, text: "Added after." }) },
  insertBeforeBlock: { operation: ({ first }) => ({ blockId: first, text: "Added before." }) },
  replaceBlock: { operation: ({ second }) => ({ blockId: second, text: "Rewritten." }) },
  deleteBlock: { operation: ({ third }) => ({ blockId: third }) },
  splitBlock: { operation: ({ first }) => ({ blockId: first, offset: 12, separator: " " }) },
  mergeBlockWithNext: { operation: ({ first }) => ({ blockId: first, separator: " " }) },
  setBlockParagraphProperties: {
    operation: ({ first }) => ({ blockId: first, properties: { alignment: "center" } }),
  },
  insertTable: {
    operation: ({ first }) => ({ blockId: first, position: "after", rows: [["A", "B"]] }),
  },
  deleteTable: { operation: ({ cell }) => ({ blockId: cell(1, 0) }) },
  commentOnBlock: { operation: ({ first }) => ({ blockId: first, comment: "Check." }) },
  insertSignatureTable: {
    operation: ({ third }) => ({ blockId: third, position: "after", parties: [{ name: "Acme" }] }),
  },
  insertTableRow: {
    operation: ({ cell }) => ({ blockId: cell(1, 0), position: "after", cellTexts: ["x", "y"] }),
  },
  deleteTableRow: { operation: ({ cell }) => ({ blockId: cell(2, 0) }) },
  insertTableColumn: {
    operation: ({ cell }) => ({
      blockId: cell(1, 0),
      position: "after",
      cellTexts: ["x", "y", "z"],
    }),
  },
  deleteTableColumn: { operation: ({ cell }) => ({ blockId: cell(1, 1) }) },
  mergeTableCells: {
    operation: ({ cell }) => ({ blockId: cell(1, 0), endBlockId: cell(2, 0) }),
  },
  splitTableCell: {
    setup: ({ cell }) => ({ type: "mergeTableCells", blockId: cell(1, 0), endBlockId: cell(2, 0) }),
    operation: ({ cell }) => ({ blockId: cell(1, 0) }),
  },
} as const satisfies Record<FolioDocumentOperationType, ReplayCase>;

const REPLAY_MODES = ["tracked-changes", "direct"] as const;

const ALL_TYPES = { suggestChanges: { operationTypes: FOLIO_DOCUMENT_OPERATION_TYPES } };

const summaryOf = (outcome: ReturnType<typeof executeFolioToolCallUntyped>) => {
  if (!outcome.ok) throw new Error(outcome.error);
  // SAFETY: suggest_changes resolves to an apply summary whenever it succeeds.
  return outcome.result as FolioAgentApplyOperationsSummary;
};

const suggest = (bridge: FolioAgentBridge, operations: readonly Record<string, unknown>[]) =>
  executeFolioToolCallUntyped("suggest_changes", { operations }, bridge, ALL_TYPES);

const blocksOf = (reviewer: FolioDocxReviewer, bridge: FolioAgentBridge): Blocks => {
  const content = reviewer.getContent();
  const byText = (text: string) =>
    content.find((block) => block.text === text)?.id ?? panic(`no block "${text}"`);
  const cells = content.filter((block) => block.table !== undefined);
  const found = executeFolioToolCallUntyped("find_text", { query: "The supplier" }, bridge);
  if (!found.ok) throw new Error(found.error);
  // SAFETY: find_text resolves to its match list whenever it succeeds.
  const [match] = (found.result as { matches: { range: unknown }[] }).matches;
  return {
    heading: byText("Agreement"),
    first: byText("The supplier delivers the goods on time."),
    second: byText("The buyer pays within thirty days."),
    third: byText("Either party may terminate on notice."),
    cell: (row, column) =>
      cells.find(({ table }) => table?.rowIndex === row && table.cellIndex === column)?.id ??
      panic(`no cell ${String(row)}:${String(column)}`),
    range: match?.range ?? panic("find_text found no range"),
  };
};

const panic = (message: string): never => {
  throw new Error(message);
};

const openSession = async (mode: (typeof REPLAY_MODES)[number]) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(fromMarkdown(MARKDOWN)), {
    author: "Agent",
  });
  return { reviewer, bridge: createReviewerBridge(reviewer, { mode }) };
};

/** Everything a replay could change: the blocks, the pending changes, the comments. */
const stateOf = (reviewer: FolioDocxReviewer) => ({
  blocks: reviewer.getContent().map(({ id, kind, text }) => ({ id, kind, text })),
  changes: reviewer.getChanges().length,
  comments: reviewer.getComments().length,
});

const cases = Object.entries(REPLAY_CASES).flatMap(([type, replayCase]) =>
  REPLAY_MODES.filter((mode) =>
    isFolioDocumentOperationModeSupported(
      // SAFETY: the keys of REPLAY_CASES are the operation types it satisfies.
      type as FolioDocumentOperationType,
      mode,
    ),
  ).map((mode) => ({ type, mode, replayCase: replayCase as ReplayCase })),
);

describe("suggest_changes replays an operation id once", () => {
  test("every operation type has a replay case in a mode that applies it", () => {
    expect(new Set(cases.map(({ type }) => type))).toEqual(new Set(FOLIO_DOCUMENT_OPERATION_TYPES));
  });

  test.each(cases)("$type, $mode", async ({ type, mode, replayCase }) => {
    const { reviewer, bridge } = await openSession(mode);
    const setup = replayCase.setup;
    if (setup !== undefined) {
      const direct = createReviewerBridge(reviewer, { mode: "direct" });
      const prepared = summaryOf(
        suggest(direct, [{ id: "setup", ...setup(blocksOf(reviewer, direct)) }]),
      );
      expect(prepared.skipped).toEqual([]);
    }
    const operation = { id: "retry-1", type, ...replayCase.operation(blocksOf(reviewer, bridge)) };

    const first = summaryOf(suggest(bridge, [operation]));
    expect(first.skipped).toEqual([]);
    expect([...first.applied, ...first.queued]).toEqual([{ id: "retry-1" }]);
    expect(first.replayed).toEqual([]);
    const after = stateOf(reviewer);

    const replay = summaryOf(suggest(bridge, [operation]));
    expect(replay.applied).toEqual([]);
    expect(replay.queued).toEqual([]);
    expect(replay.skipped).toEqual([]);
    expect(replay.replayed).toEqual([{ id: "retry-1" }]);
    expect(replay.receipts).toEqual(first.receipts);
    expect(stateOf(reviewer)).toEqual(after);
  });
});

describe("an operation id reused for a different operation", () => {
  test("fails the call and changes nothing", async () => {
    const { reviewer, bridge } = await openSession("tracked-changes");
    const { first } = blocksOf(reviewer, bridge);
    summaryOf(
      suggest(bridge, [{ id: "op-1", type: "insertAfterBlock", blockId: first, text: "A." }]),
    );
    const before = stateOf(reviewer);

    const reused = suggest(bridge, [
      { id: "op-1", type: "insertAfterBlock", blockId: first, text: "B." },
    ]);

    expect(reused.ok).toBe(false);
    expect(!reused.ok && reused.error).toContain('"op-1"');
    expect(stateOf(reviewer)).toEqual(before);
  });
});

describe("a replay beside new operations", () => {
  test("applies only the new ones and keeps receipts in input order", async () => {
    const { reviewer, bridge } = await openSession("tracked-changes");
    const { first, second } = blocksOf(reviewer, bridge);
    const once = { id: "once", type: "insertAfterBlock", blockId: first, text: "Once." };
    const original = summaryOf(suggest(bridge, [once]));

    const next = summaryOf(
      suggest(bridge, [
        { id: "new", type: "insertAfterBlock", blockId: second, text: "New." },
        once,
      ]),
    );

    expect(next.applied).toEqual([{ id: "new" }]);
    expect(next.replayed).toEqual([{ id: "once" }]);
    expect(
      next.receipts.map(({ operationId, operationIndex }) => [operationId, operationIndex]),
    ).toEqual([
      ["new", 0],
      ["once", 1],
    ]);
    expect(next.receipts.at(1)?.affected).toEqual(original.receipts.at(0)?.affected);
    expect(reviewer.getContent().filter(({ text }) => text === "Once.")).toHaveLength(1);
  });

  test("a skipped operation is not recorded, so resending it tries again", async () => {
    const { reviewer, bridge } = await openSession("tracked-changes");
    const { first } = blocksOf(reviewer, bridge);
    const stale = {
      id: "stale",
      type: "insertAfterBlock",
      blockId: first,
      text: "Later.",
      precondition: { blockTextHash: "h0" },
    };
    expect(summaryOf(suggest(bridge, [stale])).skipped.map(({ id }) => id)).toEqual(["stale"]);

    const retried = summaryOf(suggest(bridge, [stale]));

    expect(retried.replayed).toEqual([]);
    expect(retried.skipped.map(({ id }) => id)).toEqual(["stale"]);
  });

  test("the session is the document, not the bridge object", async () => {
    const { reviewer, bridge } = await openSession("tracked-changes");
    const { first } = blocksOf(reviewer, bridge);
    const operation = { id: "shared", type: "insertAfterBlock", blockId: first, text: "One." };
    summaryOf(suggest(bridge, [operation]));

    const replay = summaryOf(suggest(createReviewerBridge(reviewer), [operation]));

    expect(replay.replayed).toEqual([{ id: "shared" }]);
    expect(reviewer.getContent().filter(({ text }) => text === "One.")).toHaveLength(1);
  });

  test("a host queue receives a queued operation once", async () => {
    const { reviewer } = await openSession("tracked-changes");
    const enqueued: string[] = [];
    const queue: FolioAgentBridge = {
      snapshot: () => reviewer.snapshot(),
      applyDocumentOperations: (batch) => {
        for (const { id } of batch.operations) enqueued.push(id);
        return {
          version: 1,
          status: "queued",
          applied: [],
          queued: batch.operations.map(({ id }) => ({ id })),
          skipped: [],
          issues: [],
          receipts: getFolioDocumentOperationReceipts(
            batch.operations,
            batch.operations.map(({ id }) => ({ id })),
          ),
          undoHandle: null,
        };
      },
      getComments: () => [],
      getChanges: () => [],
      replyToComment: () => false,
      resolveComment: () => false,
    };
    const { first } = blocksOf(reviewer, createReviewerBridge(reviewer));
    const operation = { id: "queued", type: "insertAfterBlock", blockId: first, text: "Later." };

    const once = summaryOf(suggest(queue, [operation]));
    const again = summaryOf(suggest(queue, [operation]));

    expect(once.queued).toEqual([{ id: "queued" }]);
    expect(again.replayed).toEqual([{ id: "queued" }]);
    expect(again.receipts).toEqual(once.receipts);
    expect(enqueued).toEqual(["queued"]);
  });
});
