/**
 * The requested-outcome oracle (support/oracle.ts) over the batches most
 * likely to break it: collisions of two operations on one block, edits next
 * to surrogate pairs, table payloads, and edits of pending tracked changes
 * after a save. An `applied` receipt must mean the saved document is what
 * was asked; a refusal must leave it as it was.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { hashFolioAIBlockText } from "@stll/folio-core/server";
import { fromMarkdown } from "@stll/folio-core/markdown";

import {
  COLLISION_FIXTURES,
  FIXTURES,
  openReviewer,
  packDocument,
  plainDocument,
} from "../support/documents.ts";
import { assertHealthy, saveAndReopen, visibleState } from "../support/invariants.ts";
import {
  expectedFailure,
  KNOWN_FAILING_COLLISIONS,
  KNOWN_FAILING_FOLLOW_UPS,
} from "../support/known-issues.ts";
import { type Block, COLLISIONS, MODES, type Mode, type Operation } from "../support/operations.ts";
import {
  applyChecked,
  compareComments,
  compareWithModel,
  EXPECTATIONS,
  expectOperation,
  modelOf,
  OPERATION_TYPES,
  type Row,
} from "../support/oracle.ts";
import { captureLinks, comparePreservedLinks } from "../support/link-oracle.ts";
import { compareTableGeometry } from "../support/table-oracle.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const ALL_FIXTURES = { ...FIXTURES, ...COLLISION_FIXTURES };
const blocksOf = (reviewer: Reviewer): Block[] => reviewer.getContent() as Block[];

describe("the requested-outcome oracle", () => {
  test("decides, for every public operation type, what an applied receipt means", () => {
    assert.deepEqual(
      Object.keys(EXPECTATIONS).sort(),
      [...OPERATION_TYPES].sort(),
      "an operation type without an entry in EXPECTATIONS: model it, or mark it `null` (no expectation yet)",
    );
  });

  const row = (id: string, text: string, extra: Partial<Row> = {}): Row => ({
    id,
    text,
    kind: "paragraph",
    ...extra,
  });
  const before = [row("a", "First clause."), row("b", "Second clause."), row("c", "Third.")];
  const expecting = (operations: Operation[]) => {
    const model = modelOf(before);
    for (const operation of operations) expectOperation(model, operation);
    return model;
  };

  test("notices a payload the engine dropped or put on the wrong block", () => {
    const replace = { type: "replaceInBlock", blockId: "b", find: "Second", replace: "Next" };
    assert.deepEqual(
      compareWithModel(expecting([replace]), [
        row("a", "First clause."),
        row("b", "Next clause."),
        row("c", "Third."),
      ]),
      [],
    );
    // The neighbour rewritten instead.
    assert.equal(
      compareWithModel(expecting([replace, { type: "deleteBlock", blockId: "b" }]), [
        row("a", "First clause."),
        row("b", "Next clause."),
      ]).length,
      1,
    );
    // One of two cell texts dropped.
    const cells = [
      row("t1", "Cell", { table: location(0) }),
      row("t2", "Cell", { table: location(1) }),
    ];
    const model = modelOf(cells);
    expectOperation(model, {
      type: "insertTableRow",
      blockId: "t1",
      cellTexts: ["New one", "New two"],
    });
    assert.equal(compareWithModel(model, [...cells, row("n", "New one")]).length, 1);
  });

  test("notices a paragraph style the request named but the result lacks", () => {
    const restyle = {
      type: "setBlockParagraphProperties",
      blockId: "a",
      properties: { styleId: "Heading2" },
    };
    assert.match(
      compareWithModel(expecting([restyle]), before).join("\n"),
      /styleId is undefined, expected "Heading2"/u,
    );
  });

  test("notices a table merge that keeps the original grid", () => {
    const cells = [
      row("left", "Party", { table: location(0) }),
      row("right", "Role", { table: location(1) }),
    ];
    const model = modelOf(cells);
    expectOperation(model, { type: "mergeTableCells", blockId: "left", endBlockId: "right" });
    assert.deepEqual(model.tableGaps, []);
    assert.match(compareTableGeometry(model.tables, cells).join("\n"), /geometry/u);
  });

  test("requires comments while any anchored character survives a text edit", () => {
    const text = "abcdef";
    for (let anchorStart = 0; anchorStart < text.length; anchorStart += 1) {
      for (let anchorEnd = anchorStart + 1; anchorEnd <= text.length; anchorEnd += 1) {
        for (let editStart = 0; editStart < text.length; editStart += 1) {
          for (let editEnd = editStart + 1; editEnd <= text.length; editEnd += 1) {
            const model = modelOf([row("b", text)]);
            expectOperation(model, {
              type: "replaceRange",
              range: { blockId: "b", startOffset: editStart, endOffset: editEnd },
              replace: "",
            });
            const comment = {
              id: 1,
              text: "review",
              anchor: text.slice(anchorStart, anchorEnd),
              blockId: "b",
            };
            const removed = editStart <= anchorStart && editEnd >= anchorEnd;
            const problems = compareComments(model, [comment], []);
            assert.equal(
              problems.length === 0,
              removed,
              JSON.stringify({ anchorStart, anchorEnd, editStart, editEnd }),
            );
            assert.equal(
              compareComments(model, [comment], [comment]).length === 0,
              !removed,
              JSON.stringify({ anchorStart, anchorEnd, editStart, editEnd }),
            );
          }
        }
      }
    }
  });

  test("preserves a comment when a split or a different block changes its quoted text", () => {
    const comment = { id: 1, text: "review", anchor: "bcd", blockId: "b" };
    const split = modelOf([row("b", "abcdef")]);
    expectOperation(split, { type: "splitBlock", blockId: "b", offset: 3 });
    assert.equal(compareComments(split, [comment], []).length, 1);
    const elsewhere = modelOf([row("a", "bcd"), row("b", "abcdef")]);
    expectOperation(elsewhere, {
      type: "replaceRange",
      range: { blockId: "a", startOffset: 0, endOffset: 3 },
      replace: "",
    });
    assert.equal(compareComments(elsewhere, [comment], []).length, 1);
  });

  test("does not guess which repeated quote a comment anchored", () => {
    const comment = { id: 1, text: "review", anchor: "foo", blockId: "b" };
    const model = modelOf([row("b", "foo foo")]);
    expectOperation(model, {
      type: "replaceRange",
      range: { blockId: "b", startOffset: 0, endOffset: 3 },
      replace: "",
    });
    assert.deepEqual(compareComments(model, [comment], []), []);
    assert.deepEqual(compareComments(model, [comment], [comment]), []);
  });

  test("a nonempty replacement carries a comment, while an empty one removes it", () => {
    const comment = { id: 1, text: "review", anchor: "Signed", blockId: "b" };
    const replaced = modelOf([row("b", "Signed")]);
    expectOperation(replaced, { type: "replaceBlock", blockId: "b", text: "Written" });
    assert.deepEqual(compareComments(replaced, [comment], [{ ...comment, anchor: "Written" }]), []);
    assert.match(
      compareComments(replaced, [comment], [{ ...comment, anchor: "Unrelated" }]).join("\n"),
      /no comment/u,
    );
    assert.equal(compareComments(replaced, [comment], []).length, 1);

    const emptied = modelOf([row("b", "Signed")]);
    expectOperation(emptied, { type: "replaceBlock", blockId: "b", text: "" });
    assert.deepEqual(compareComments(emptied, [comment], []), []);
    assert.equal(compareComments(emptied, [comment], [comment]).length, 1);
  });

  test("a deleted comment cannot satisfy a new comment with the same text", () => {
    const old = { id: 1, text: "same note", anchor: "Old", blockId: "a" };
    const model = modelOf([row("a", "Old"), row("b", "New")]);
    expectOperation(model, { type: "deleteBlock", blockId: "a" });
    expectOperation(model, {
      type: "commentOnBlock",
      blockId: "b",
      comment: { text: "same note" },
    });
    assert.deepEqual(
      compareComments(model, [old], [{ id: 2, text: "same note", anchor: "New", blockId: "b" }]),
      [],
    );
    const problems = compareComments(model, [old], [{ ...old, anchor: "New", blockId: "b" }]);
    assert.match(problems.join("\n"), /no comment/u);
    assert.match(problems.join("\n"), /outlived the block it anchored/u);
  });

  test("uses live anchors and ids when an accepted pending join moves one of three equal comments", () => {
    const a = { id: 1, text: "same note", anchor: "First", blockId: "a" };
    const moved = { id: 2, text: "same note", anchor: "Signed", blockId: "c" };
    const c = { id: 3, text: "same note", anchor: "Last", blockId: "c" };
    const model = modelOf([row("a", "First"), row("b", "Signed"), row("c", "Signed Last")]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    const live = [a, { ...moved, blockId: "b" }, c];
    assert.deepEqual(compareComments(model, [a, moved, c], [a, c], live), []);
    assert.match(
      compareComments(model, [a, moved, c], [a, moved, c], live).join("\n"),
      /outlived the block it anchored/u,
    );
  });

  test("a comment on a pending join survives an edit to another block", () => {
    const joined = { id: 1, text: "review", anchor: "Signed", blockId: "b" };
    const model = modelOf(
      [row("a", "First Signed"), row("c", "Other")],
      [row("a", "First"), row("b", "Signed"), row("c", "Other")],
    );
    expectOperation(model, { type: "replaceBlock", blockId: "c", text: "Revised" });
    assert.deepEqual(compareComments(model, [joined], [joined], [joined]), []);
    assert.equal(compareComments(model, [joined], [], [joined]).length, 1);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [joined], [], [joined]), []);
    assert.match(
      compareComments(model, [joined], [joined], [joined]).join("\n"),
      /outlived the block it anchored/u,
    );
  });

  test("an explicit block deletion removes a live comment even when accepted text differs", () => {
    const comment = { id: 1, text: "review", anchor: "Live quote", blockId: "b" };
    const model = modelOf([row("b", "Accepted text")], [row("b", "Live quote")]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [comment], [], [comment]), []);
    assert.match(
      compareComments(model, [comment], [comment], [comment]).join("\n"),
      /outlived the block it anchored/u,
    );
  });

  test("deleting the first block of a spanning comment keeps its surviving anchor", () => {
    const comment = { id: 1, text: "review", anchor: "FirstSecond", blockId: "b" };
    const model = modelOf([row("b", "First"), row("c", "Second")]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [comment], [comment]), []);
    assert.equal(compareComments(model, [comment], []).length, 1);
  });

  test("a suggested block deletion leaves the live comment in place", () => {
    const comment = { id: 1, text: "review", anchor: "Signed", blockId: "b" };
    const model = modelOf([row("b", "Signed")]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [comment], [comment], [comment], "suggested"), []);
  });
});

test("a style edit preserves the target of an unchanged link", async () => {
  const reviewer = await openReviewer(
    await packDocument(
      fromMarkdown("# Terms\n\nRead the [schedule](https://example.com/schedule)."),
    ),
  );
  const block = blocksOf(reviewer).find((candidate) => candidate.text.includes("schedule"));
  assert.ok(block);
  assert.ok(captureLinks(reviewer).get(block.id)?.targets.includes("https://example.com/schedule"));
  const { applied } = await applyChecked(
    reviewer,
    [
      {
        type: "setBlockParagraphProperties",
        blockId: block.id,
        properties: { alignment: "center" },
      },
    ],
    "direct",
    "link after style edit",
  );
  assert.equal(applied.length, 1);
});

test("a missing saved link projection fails for an unchanged reader block", async () => {
  const reviewer = await openReviewer(
    await packDocument(fromMarkdown("Read the [schedule](https://example.com/schedule).")),
  );
  const before = captureLinks(reviewer);
  const { reopened } = await saveAndReopen(reviewer, "link projection fixture");
  const after = new Map(captureLinks(reopened));
  const linked = [...before.keys()].at(0);
  assert.ok(linked);
  assert.ok(after.delete(linked));

  const unchanged = comparePreservedLinks({ before, after, afterRows: blocksOf(reopened) });
  assert.equal(unchanged.checked, 1);
  assert.match(unchanged.problems.join("\n"), /has no matching link projection/u);

  const editedRows = blocksOf(reopened).map(({ id, text }) => ({
    id,
    text: id === linked ? `${text} changed` : text,
  }));
  assert.deepEqual(comparePreservedLinks({ before, after, afterRows: editedRows }), {
    problems: [],
    checked: 0,
  });
});

test("a signature table has one cell per party after save and reopen", async () => {
  const reviewer = await openReviewer(await plainDocument());
  const block = blocksOf(reviewer).find((candidate) => candidate.text === "Signed in two copies.");
  assert.ok(block);
  const { applied } = await applyChecked(
    reviewer,
    [
      {
        type: "insertSignatureTable",
        blockId: block.id,
        parties: [
          { name: "Supplier", signatory: "A. Example" },
          { name: "Buyer", title: "Director" },
        ],
      },
    ],
    "direct",
    "signature table geometry",
  );
  assert.equal(applied.length, 1);
});

test("cell merge and split predict the saved grid", async () => {
  const reviewer = await openReviewer(await ALL_FIXTURES.tables());
  const left = blocksOf(reviewer).find((candidate) => candidate.text === "Item");
  const right = blocksOf(reviewer).find((candidate) => candidate.text === "Price");
  assert.ok(left);
  assert.ok(right);
  const merged = await applyChecked(
    reviewer,
    [{ type: "mergeTableCells", blockId: left.id, endBlockId: right.id }],
    "direct",
    "merge grid",
  );
  assert.equal(merged.applied.length, 1);
  const split = await applyChecked(
    reviewer,
    [{ type: "splitTableCell", blockId: left.id }],
    "direct",
    "split grid",
  );
  assert.equal(split.applied.length, 1);
});

test("a deleted merge neighbor does not join the following table cell", async () => {
  const reviewer = await openReviewer(await ALL_FIXTURES.tables());
  const heading = blocksOf(reviewer).find((candidate) => candidate.text === "Price Schedule");
  const neighbor = blocksOf(reviewer).find(
    (candidate) => candidate.text === "The prices below apply.",
  );
  assert.ok(heading);
  assert.ok(neighbor);
  const { applied } = await applyChecked(
    reviewer,
    [
      { type: "mergeBlockWithNext", blockId: heading.id, separator: " " },
      { type: "deleteBlock", blockId: neighbor.id },
    ],
    "tracked-changes",
    "join then remove neighbor",
  );
  assert.equal(applied.length, 2);
});

function location(cellIndex: number) {
  return {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex: 0,
    cellIndex,
    gridColumnIndex: cellIndex,
    columnSpan: 1,
    rowSpan: 1,
  };
}

/**
 * `docxToMarkdown` writes a table with merged cells as HTML, which the reader
 * comparison does not parse; that fixture is saved and reopened only.
 */
const READERS_COMPARABLE = (fixture: string) => fixture !== "mergedTable";

const persistedAfterBulkAcceptance = async (reviewer: Reviewer) => {
  const saved = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  saved.acceptAll();
  return visibleState(saved);
};

/** Run one collision on a fresh reviewer; what applied, or null when none fits. */
const runCollision = async (
  fixture: string,
  bytes: Uint8Array,
  mode: Mode,
  name: string,
): Promise<string | null> => {
  const context = `${fixture} / ${mode}: ${name}`;
  const reviewer = await openReviewer(bytes);
  const persisted =
    mode === "suggested" ? { persisted: await persistedAfterBulkAcceptance(reviewer) } : {};
  const operations = COLLISIONS[name]?.(blocksOf(reviewer));
  if (!operations) return null;
  const { applied, issues } = await applyChecked(reviewer, operations, mode, context);
  if (mode === "suggested") reviewer.acceptAll();
  if (READERS_COMPARABLE(fixture)) await assertHealthy(reviewer, context, persisted);
  else await saveAndReopen(reviewer, context, persisted);
  return `${name}: applied ${applied.length}${issues.length > 0 ? `, refused ${issues.join(", ")}` : ""}`;
};

describe("collision batches", () => {
  for (const [fixture, build] of Object.entries(ALL_FIXTURES)) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_COLLISIONS.filter(
        (entry) => entry.fixture === fixture && entry.mode === mode,
      );
      test(`${fixture} / ${mode}: every collision does what its applied operations asked`, async () => {
        const bytes = await build();
        const ran: string[] = [];
        for (const name of Object.keys(COLLISIONS)) {
          if (known.some((entry) => entry.collision === name)) continue;
          const outcome = await runCollision(fixture, bytes, mode, name);
          if (outcome !== null) ran.push(outcome);
        }
        if (process.env["FOLIO_ORACLE_GAPS"])
          console.log(`${fixture} / ${mode}:\n  ${ran.join("\n  ")}`);
        assert.ok(ran.length > 0, "no collision fits this fixture");
      });
      for (const { collision, finding, symptom } of known) {
        expectedFailure(finding, `${fixture} / ${mode}: ${collision}`, symptom, async () => {
          await runCollision(fixture, await build(), mode, collision);
        });
      }
    }
  }
});

/**
 * A first edit, saved and reopened, then a second one that names what the
 * first left pending.
 */
type FollowUp = {
  first: (blocks: readonly Block[]) => Operation[];
  second: (blocks: readonly Block[]) => Operation[] | null;
  /** The issue code every mode must refuse the second edit with. */
  refusedWith?: string;
};

const INSERTED = "An inserted clause about delivery terms.";
const anchorOf = (blocks: readonly Block[]) =>
  blocks.find((block) => block.text.startsWith("The Supplier")) as Block;
const insertedOf = (blocks: readonly Block[]) => blocks.find((block) => block.text === INSERTED);
const insertFirst = (blocks: readonly Block[]): Operation[] => [
  { type: "insertAfterBlock", blockId: anchorOf(blocks).id, text: INSERTED },
];
const onInserted =
  (build: (block: Block) => Operation) =>
  (blocks: readonly Block[]): Operation[] | null => {
    const block = insertedOf(blocks);
    return block ? [build(block)] : null;
  };
const rangeOver = (block: Block, word: string) => {
  const start = block.text.indexOf(word);
  return createRange(block, start, start + word.length);
};

const FOLLOW_UPS: Record<string, FollowUp> = {
  replaceInPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "replaceInBlock",
      blockId: block.id,
      find: "delivery",
      replace: "payment",
    })),
  },
  replaceRangeInPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "replaceRange",
      range: rangeOver(block, "clause"),
      replace: "term",
    })),
  },
  deletePendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({ type: "deleteBlock", blockId: block.id })),
  },
  formatPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "formatRange",
      range: rangeOver(block, "inserted"),
      formatting: { bold: true },
    })),
  },
  splitPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "splitBlock",
      blockId: block.id,
      offset: block.text.indexOf("about"),
    })),
  },
  replaceOverPendingReplacement: {
    first: (blocks) => [
      { type: "replaceInBlock", blockId: anchorOf(blocks).id, find: "goods", replace: "products" },
    ],
    second: (blocks) => {
      const block = blocks.find((candidate) => candidate.text.includes("products"));
      return block
        ? [{ type: "replaceRange", range: rangeOver(block, "products"), replace: "wares" }]
        : null;
    },
  },
  rewritePendingReplacement: {
    first: (blocks) => [
      { type: "replaceInBlock", blockId: anchorOf(blocks).id, find: "goods", replace: "products" },
    ],
    second: (blocks) => {
      const block = blocks.find((candidate) => candidate.text.includes("products"));
      return block
        ? [{ type: "replaceBlock", blockId: block.id, text: "The Supplier delivers promptly." }]
        : null;
    },
  },
  // A paragraph pending deletion reads as a blank block, which a model may name.
  // Text typed there stays joined to the next paragraph once accepted, so
  // rewriting it is refused rather than glued onto that paragraph.
  rewritePendingDeletion: {
    first: deleteFirst,
    second: onDeleted((block) => ({ type: "replaceBlock", blockId: block.id, text: "Rewritten." })),
    refusedWith: "pendingDeletion",
  },
  insertAfterPendingDeletion: {
    first: deleteFirst,
    second: onDeleted((block) => ({ type: "insertAfterBlock", blockId: block.id, text: "Added." })),
  },
  insertBeforePendingDeletion: {
    first: deleteFirst,
    second: onDeleted((block) => ({
      type: "insertBeforeBlock",
      blockId: block.id,
      text: "Added.",
    })),
  },
};

function deleteFirst(blocks: readonly Block[]): Operation[] {
  return [{ type: "deleteBlock", blockId: anchorOf(blocks).id }];
}

function onDeleted(build: (block: Block) => Operation) {
  return (blocks: readonly Block[]): Operation[] | null => {
    const block = blocks.find((candidate) => candidate.text === "");
    return block ? [build(block)] : null;
  };
}

function createRange(block: Block, start: number, end: number) {
  return {
    type: "textRange",
    story: "main",
    blockId: block.id,
    startOffset: start,
    endOffset: end,
    selectedTextHash: hashFolioAIBlockText(block.text.slice(start, end)),
  };
}

const runFollowUp = async (name: string, mode: Mode): Promise<void> => {
  const followUp = FOLLOW_UPS[name] as FollowUp;
  const context = `tracked, saved, then ${mode}: ${name}`;
  const reviewer = await openReviewer(await plainDocument());
  const first = await applyChecked(
    reviewer,
    followUp.first(blocksOf(reviewer)),
    "tracked-changes",
    context,
  );
  assert.ok(first.applied.length > 0, `${context}: the first edit was refused: ${first.issues}`);
  const reopened = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  const persisted =
    mode === "suggested" ? { persisted: await persistedAfterBulkAcceptance(reopened) } : {};
  const operations = followUp.second(blocksOf(reopened));
  assert.ok(operations, `${context}: the second edit found nothing to name`);
  const second = await applyChecked(reopened, operations, mode, context);
  if (followUp.refusedWith !== undefined) {
    assert.deepEqual(second.applied, [], `${context}: applied what it must refuse`);
    assert.ok(
      second.issues.every((issue) => issue.endsWith(`: ${followUp.refusedWith}`)),
      `${context}: refused with ${second.issues.join(", ")}, not ${followUp.refusedWith}`,
    );
  }
  if (mode === "suggested") reopened.acceptAll();
  await assertHealthy(reopened, context, persisted);
};

describe("edits of what a saved tracked edit left pending", () => {
  for (const name of Object.keys(FOLLOW_UPS)) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_FOLLOW_UPS.find(
        (entry) => entry.followUp === name && entry.mode === mode,
      );
      const title = `${name} (${mode})`;
      if (known) {
        expectedFailure(known.finding, title, known.symptom, () => runFollowUp(name, mode));
      } else {
        test(title, () => runFollowUp(name, mode));
      }
    }
  }
});

/** Style ids the package does not define as paragraph styles, named by an operation. */
const STYLE_REQUESTS: Record<string, (block: Block) => Operation> = {
  "insertAfterBlock with an undefined style": (block) => ({
    type: "insertAfterBlock",
    blockId: block.id,
    text: "Styled clause.",
    styleId: "NoSuchStyle",
  }),
  "replaceBlock with an undefined style": (block) => ({
    type: "replaceBlock",
    blockId: block.id,
    text: "Styled clause.",
    styleId: "NoSuchStyle",
  }),
  "setBlockParagraphProperties with an undefined style": (block) => ({
    type: "setBlockParagraphProperties",
    blockId: block.id,
    properties: { styleId: "NoSuchStyle" },
  }),
  "setBlockParagraphProperties with a table style": (block) => ({
    type: "setBlockParagraphProperties",
    blockId: block.id,
    properties: { styleId: "TableGrid" },
  }),
};

const runStyleRequest = async (name: string, mode: Mode): Promise<void> => {
  const reviewer = await openReviewer(await ALL_FIXTURES.tables());
  const block = blocksOf(reviewer).find((candidate) => candidate.text === "Taxes are extra.");
  assert.ok(block);
  const build = STYLE_REQUESTS[name] as (block: Block) => Operation;
  await applyChecked(reviewer, [build(block)], mode, `${name} (${mode})`);
};

describe("a style an operation names is a paragraph style of the saved package, or the operation is refused", () => {
  for (const name of Object.keys(STYLE_REQUESTS)) {
    for (const mode of ["direct", "tracked-changes"] as const) {
      test(`${name} (${mode})`, () => runStyleRequest(name, mode));
    }
  }
});

describe("findings the oracle pins", () => {
  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`two paragraph-property operations on one block in one batch: the later one is refused (${mode})`, async () => {
      const reviewer = await openReviewer(await plainDocument());
      const block = blocksOf(reviewer).at(-1) as Block;
      const align = (alignment: string): Operation => ({
        type: "setBlockParagraphProperties",
        blockId: block.id,
        properties: { alignment },
      });
      const { applied, issues } = await applyChecked(
        reviewer,
        [align("center"), align("right")],
        mode,
        `alignment twice (${mode})`,
      );
      assert.deepEqual(applied, ["op-1"]);
      assert.deepEqual(issues, ["op-2: overlappingOperation"]);
    });
  }

  for (const mode of ["tracked-changes", "suggested"] as const) {
    test(`a row deletion over a pending word leaves the reader no text of that row (${mode})`, async () => {
      const reviewer = await openReviewer(await ALL_FIXTURES.tables());
      const cell = blocksOf(reviewer).find((block) => block.text === "Price") as Block;
      await applyChecked(
        reviewer,
        [{ type: "replaceInBlock", blockId: cell.id, find: "Price", replace: "amended" }],
        mode,
        "a pending word",
      );
      await applyChecked(
        reviewer,
        [{ type: "deleteTableRow", blockId: cell.id }],
        mode,
        "then its row's deletion",
      );
      assert.ok(
        !blocksOf(reviewer).some((block) => block.text.includes("amended")),
        "the reader still lists the deleted row's pending word",
      );
    });
  }
});
