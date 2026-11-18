import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";

import { FIXTURES, openReviewer, packDocument } from "../support/documents.ts";
import {
  assertRequestedOutcome,
  capture,
  compareComments,
  expectOperation,
  modelOf,
  rowsOf,
} from "../support/oracle.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Operation = Parameters<
  Reviewer["applyDocumentOperations"]
>[0]["operations"][number];

const applyChecked = async (reviewer: Reviewer, operation: Operation) => {
  const pre = await capture(reviewer, "tracked-changes");
  const input = {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [operation],
  } as const;
  const receipt = reviewer.applyDocumentOperations(
    JSON.parse(JSON.stringify(input)),
    { undefinedReferences: "refuse" },
  );
  assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
  assert.deepEqual(receipt.skipped, []);
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  model.mode = pre.mode;
  expectOperation(model, operation);
  assert.deepEqual(model.unmodelled, []);
  await assertRequestedOutcome(
    reviewer,
    pre,
    { applied: [operation] },
    operation.id,
  );
  return { pre, model, after: await capture(reviewer, "tracked-changes") };
};

const pending = async (right = "token right") => {
  const reviewer = await openReviewer(
    await packDocument(
      fromMarkdown(`token left\n\n${right}\n\nUnrelated token survivor.`),
    ),
  );
  const [first, second, unrelated] = rowsOf(reviewer);
  assert.ok(first && second && unrelated);
  const ids = [first.id, second.id, unrelated.id];
  await applyChecked(reviewer, {
    id: "join",
    type: "mergeBlockWithNext",
    blockId: first.id,
    separator: " ",
  });
  const pre = await capture(reviewer, "tracked-changes");
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  const ownership = [...model.pendingJoins.values()].at(0);
  assert.ok(ownership);
  assert.equal(
    ownership.sources.map(({ text }) => text).join("") +
      ownership.destination.text,
    pre.rows.at(0)?.text,
  );
  assert.equal(pre.rows.at(0)?.id, second.id);
  return { reviewer, first, second, unrelated, ids };
};

const rangeOf = (reviewer: Reviewer, blockId: string, start = 0, end = 5) => {
  const block = rowsOf(reviewer).find(({ id }) => id === blockId);
  assert.ok(block);
  const range = createFolioAITextRangeHandle({
    blockId,
    text: block.text,
    startOffset: start,
    endOffset: end,
  });
  assert.ok(range);
  return range;
};

for (const side of ["source", "destination"] as const)
  for (const kind of [
    "replaceInBlock",
    "replaceRange",
    "formatRange",
    "commentOnRange",
    "commentOnBlockQuote",
    "commentOnBlockWhole",
  ] as const)
    test(`${side} ${kind} uses its live slice of the accepted joined paragraph`, async () => {
      const { reviewer, first, second, ids } = await pending();
      const blockId = side === "source" ? first.id : second.id;
      const range = rangeOf(reviewer, blockId);
      const operations = {
        replaceInBlock: {
          id: "replace-word",
          type: "replaceInBlock",
          blockId,
          find: "token",
          replace: "EDIT",
        },
        replaceRange: {
          id: "replace-range",
          type: "replaceRange",
          range,
          replace: "EDIT",
        },
        formatRange: {
          id: "format-range",
          type: "formatRange",
          range,
          formatting: { italic: true },
        },
        commentOnRange: {
          id: "comment-range",
          type: "commentOnRange",
          range,
          comment: { text: "Review range." },
        },
        commentOnBlockQuote: {
          id: "comment-quote",
          type: "commentOnBlock",
          blockId,
          quote: "token",
          comment: { text: "Review quote." },
        },
        commentOnBlockWhole: {
          id: "comment-block",
          type: "commentOnBlock",
          blockId,
          comment: { text: "Review block." },
        },
      } as const satisfies Record<typeof kind, Operation>;
      await applyChecked(reviewer, operations[kind]);
      assert.deepEqual(
        rowsOf(reviewer).map(({ id }) => id),
        ids,
        "Text edits preserve the live source paragraph identities",
      );
    });

test("the saved 1842 sequence retires the comment whose joined destination text is removed", async () => {
  const reviewer = await openReviewer(await FIXTURES.tables());
  const operations = [
    {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: "412DDC7F",
      separator: " ",
    },
    {
      id: "comment",
      type: "commentOnRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: "3CEFFC83",
        startOffset: 11,
        endOffset: 16,
        selectedTextHash: "h47ar1c",
      },
      comment: { text: "Warranty invoice buyer goods term period term buyer." },
    },
    {
      id: "remove",
      type: "replaceInBlock",
      blockId: "3CEFFC83",
      find: "below",
      replace: "",
      precondition: { blockTextHash: "h1habca" },
    },
  ] as const satisfies readonly Operation[];
  for (const operation of operations) await applyChecked(reviewer, operation);
  const after = await capture(reviewer, "tracked-changes");
  assert.deepEqual(after.comments, []);
  assert.equal(after.rows.at(0)?.id, "3CEFFC83");
  assert.equal(after.rows.at(0)?.text, "Price Schedule The prices  apply.");
});

test("joined coordinates distinguish equal quotes in different live paragraphs", async () => {
  const { reviewer, first, second } = await pending();
  await applyChecked(reviewer, {
    id: "comment",
    type: "commentOnRange",
    range: rangeOf(reviewer, second.id),
    comment: { text: "Destination token." },
  });
  const result = await applyChecked(reviewer, {
    id: "delete-source-token",
    type: "replaceInBlock",
    blockId: first.id,
    find: "token",
    replace: "",
  });
  const [comment] = result.after.comments;
  assert.ok(comment);
  assert.equal(comment.anchor, "token");
  assert.deepEqual(
    compareComments(
      result.model,
      result.pre.comments,
      [],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).length,
    1,
    "An unrelated surviving destination comment cannot disappear",
  );
  assert.equal(
    compareComments(
      result.model,
      result.pre.comments,
      [{ ...comment, anchor: "drift" }],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).length,
    1,
    "A preserved joined quote cannot drift",
  );
  const unrelated = result.after.rows.at(-1);
  assert.ok(unrelated);
  assert.equal(
    compareComments(
      result.model,
      result.pre.comments,
      [{ ...comment, blockId: unrelated.id }],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).length,
    1,
    "A preserved joined comment cannot move to another paragraph",
  );
});

test("a quote in the other join component cannot excuse retention of a fully removed anchor", async () => {
  const { reviewer, second } = await pending();
  await applyChecked(reviewer, {
    id: "comment",
    type: "commentOnRange",
    range: rangeOf(reviewer, second.id),
    comment: { text: "Remove destination token." },
  });
  const result = await applyChecked(reviewer, {
    id: "delete-destination-token",
    type: "replaceInBlock",
    blockId: second.id,
    find: "token",
    replace: "",
  });
  assert.deepEqual(result.after.comments, []);
  const stale = result.pre.comments.at(0);
  assert.ok(stale);
  assert.match(
    compareComments(
      result.model,
      result.pre.comments,
      [stale],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).join("\n"),
    /outlived/u,
  );
});

test("an unproven join keeps preservation checks and structural target refusals", async () => {
  const { reviewer, second } = await pending();
  await applyChecked(reviewer, {
    id: "comment",
    type: "commentOnRange",
    range: rangeOf(reviewer, second.id),
    comment: { text: "Keep when mapping is unknown." },
  });
  const pre = await capture(reviewer, "tracked-changes");
  const wrong = pre.rows.map((row) =>
    row.id === second.id
      ? Object.assign({}, row, { text: `not reconstructed ${row.text}` })
      : row,
  );
  const model = modelOf(wrong, pre.liveRows, pre.joinSources);
  expectOperation(model, {
    type: "replaceInBlock",
    blockId: second.id,
    find: "token",
    replace: "",
  });
  assert.equal(model.unmodelled.length, 1);
  assert.equal(
    compareComments(
      model,
      pre.comments,
      [],
      pre.rows,
      pre.liveComments,
      "tracked-changes",
    ).length,
    1,
    "Unknown mapping must not excuse known comment loss",
  );
  const structural = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  expectOperation(structural, {
    type: "splitBlock",
    blockId: second.id,
    offset: 1,
  });
  assert.match(structural.unmodelled.join("\n"), /pending join/u);
});

test("repeated quotes within one live component retain genuine location uncertainty", async () => {
  const { reviewer, second } = await pending("token token");
  await applyChecked(reviewer, {
    id: "comment",
    type: "commentOnRange",
    range: rangeOf(reviewer, second.id),
    comment: { text: "An unspecified token." },
  });
  const pre = await capture(reviewer, "tracked-changes");
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  expectOperation(model, {
    type: "replaceInBlock",
    blockId: second.id,
    find: "token",
    replace: "",
  });
  assert.deepEqual(model.unmodelled, []);
  assert.deepEqual(
    compareComments(
      model,
      pre.comments,
      [],
      pre.rows,
      pre.liveComments,
      "tracked-changes",
    ),
    [],
    "The read comment quote alone cannot identify either identical placement",
  );
});

test("a source comment keeps its accepted quote after its locator moves to the joined destination", async () => {
  const { reviewer, first, second } = await pending();
  await applyChecked(reviewer, {
    id: "comment-source",
    type: "commentOnRange",
    range: rangeOf(reviewer, first.id, 6, 10),
    comment: { text: "Keep the source word." },
  });
  const result = await applyChecked(reviewer, {
    id: "delete-destination-token",
    type: "replaceInBlock",
    blockId: second.id,
    find: "token",
    replace: "",
  });
  const accepted = result.pre.comments.at(0);
  const live = result.pre.liveComments.at(0);
  const after = result.after.comments.at(0);
  assert.ok(accepted && live && after);
  assert.equal(accepted.blockId, second.id);
  assert.equal(live.blockId, first.id);
  assert.equal(accepted.anchor, "left");
  assert.equal(after.anchor, "left");
  assert.equal(
    compareComments(
      result.model,
      result.pre.comments,
      [{ ...after, anchor: "drift" }],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).length,
    1,
    "A relocated source comment cannot lose its known accepted quote",
  );
});

test("a relocated source comment compares the accepted quote without pending deleted words", async () => {
  const { reviewer, first, second } = await pending();
  await applyChecked(reviewer, {
    id: "comment-source",
    type: "commentOnRange",
    range: rangeOf(reviewer, first.id, 0, 10),
    comment: { text: "Review the source phrase." },
  });
  await applyChecked(reviewer, {
    id: "delete-source-token",
    type: "replaceInBlock",
    blockId: first.id,
    find: "token",
    replace: "",
  });
  const result = await applyChecked(reviewer, {
    id: "delete-destination-token",
    type: "replaceInBlock",
    blockId: second.id,
    find: "token",
    replace: "",
  });
  const accepted = result.pre.comments.at(0);
  const live = result.pre.liveComments.at(0);
  const after = result.after.comments.at(0);
  assert.ok(accepted && live && after);
  assert.equal(accepted.blockId, second.id);
  assert.equal(live.blockId, first.id);
  assert.equal(accepted.anchor, " left");
  assert.equal(live.anchor, "token left");
  assert.equal(after.anchor, accepted.anchor);
  assert.equal(
    compareComments(
      result.model,
      result.pre.comments,
      [{ ...after, anchor: live.anchor }],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).length,
    1,
    "The live quote cannot replace the expected accepted quote",
  );
});

test("removing a relocated comment's remaining accepted quote retires its pending live range", async () => {
  const { reviewer, first } = await pending();
  await applyChecked(reviewer, {
    id: "comment-source",
    type: "commentOnRange",
    range: rangeOf(reviewer, first.id, 0, 10),
    comment: { text: "Review the remaining source phrase." },
  });
  await applyChecked(reviewer, {
    id: "delete-source-token",
    type: "replaceInBlock",
    blockId: first.id,
    find: "token",
    replace: "",
  });
  const result = await applyChecked(reviewer, {
    id: "delete-source-rest",
    type: "replaceInBlock",
    blockId: first.id,
    find: " left",
    replace: "",
  });
  assert.deepEqual(result.after.comments, []);
  const accepted = result.pre.comments.at(0);
  const live = result.pre.liveComments.at(0);
  assert.ok(accepted && live);
  assert.equal(accepted.anchor, " left");
  assert.equal(live.anchor, "token left");
  assert.match(
    compareComments(
      result.model,
      result.pre.comments,
      [accepted],
      result.after.rows,
      result.pre.liveComments,
      "tracked-changes",
    ).join("\n"),
    /outlived/u,
  );
});

const applyUnmodelledStructural = async (
  reviewer: Reviewer,
  operation: Operation,
) => {
  const pre = await capture(reviewer, "tracked-changes");
  const receipt = reviewer.applyDocumentOperations(
    JSON.parse(
      JSON.stringify({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [operation],
      }),
    ),
    { undefinedReferences: "refuse" },
  );
  assert.equal(receipt.applied.length, 1, JSON.stringify(receipt));
  assert.deepEqual(receipt.skipped, []);
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  model.mode = pre.mode;
  expectOperation(model, operation);
  assert.equal(model.unmodelled.length, 1);
  assert.match(
    model.unmodelled.at(0) ?? "",
    new RegExp(`^${operation.type}:`, "u"),
  );
  assert.equal(
    [...model.pendingJoins.values()].filter(
      ({ destinationPrediction }) => destinationPrediction.type === "unknown",
    ).length,
    1,
  );
  await assertRequestedOutcome(
    reviewer,
    pre,
    { applied: [operation] },
    operation.id,
  );
  return { pre, model, after: await capture(reviewer, "tracked-changes") };
};

for (const kind of [
  "splitBlock",
  "mergeBlockWithNext",
  "replaceBlock",
] as const)
  test(`an applied unmodelled ${kind} expires only its joined destination prediction`, async () => {
    const { reviewer, second } = await pending();
    await applyChecked(reviewer, {
      id: "comment-before-structure",
      type: "commentOnRange",
      range: rangeOf(reviewer, second.id, 6, 11),
      comment: { text: "Keep the structurally moved comment." },
    });
    const operations = {
      splitBlock: {
        id: "split",
        type: "splitBlock",
        blockId: second.id,
        offset: 6,
      },
      mergeBlockWithNext: {
        id: "merge-next",
        type: "mergeBlockWithNext",
        blockId: second.id,
        separator: " ",
      },
      replaceBlock: {
        id: "rewrite",
        type: "replaceBlock",
        blockId: second.id,
        text: "x",
      },
    } as const satisfies Record<typeof kind, Operation>;
    const result = await applyUnmodelledStructural(reviewer, operations[kind]);
    const before = result.pre.comments.at(0);
    const after = result.after.comments.at(0);
    assert.ok(before && after);
    assert.equal(after.id, before.id);
    assert.ok(
      after.blockId !== before.blockId || after.anchor !== before.anchor,
      "The actual producer changes the previously predicted owner or quote",
    );
    assert.equal(
      compareComments(
        result.model,
        result.pre.comments,
        [],
        result.after.rows,
        result.pre.liveComments,
        "tracked-changes",
      ).length,
      1,
      "Unmodelled structure cannot excuse losing the comment",
    );
  });

test("an unmodelled effect in one join cannot weaken another join's owner and quote proof", async () => {
  const reviewer = await openReviewer(
    await packDocument(
      fromMarkdown("left one\n\nright one\n\nleft two\n\nright two"),
    ),
  );
  const [a, b, c, d] = rowsOf(reviewer);
  assert.ok(a && b && c && d);
  await applyChecked(reviewer, {
    id: "join-one",
    type: "mergeBlockWithNext",
    blockId: a.id,
    separator: " ",
  });
  await applyChecked(reviewer, {
    id: "join-two",
    type: "mergeBlockWithNext",
    blockId: c.id,
    separator: " ",
  });
  await applyChecked(reviewer, {
    id: "comment-two",
    type: "commentOnRange",
    range: rangeOf(reviewer, d.id, 0, 5),
    comment: { text: "Keep the untouched second join." },
  });
  const result = await applyUnmodelledStructural(reviewer, {
    id: "rewrite-one",
    type: "replaceBlock",
    blockId: b.id,
    text: "x",
  });
  const comment = result.after.comments.at(0);
  assert.ok(comment);
  assert.equal(comment.blockId, d.id);
  assert.equal(comment.anchor, "right");
  for (const changed of [
    { ...comment, anchor: "drift" },
    { ...comment, blockId: b.id },
  ])
    assert.equal(
      compareComments(
        result.model,
        result.pre.comments,
        [changed],
        result.after.rows,
        result.pre.liveComments,
        "tracked-changes",
      ).length,
      1,
      "The unrelated join retains its exact owner/quote checks",
    );
});
