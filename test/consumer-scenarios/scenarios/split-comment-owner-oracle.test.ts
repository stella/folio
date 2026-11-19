import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { openReviewer } from "../support/documents.ts";
import { parseFlowFile } from "../support/flow-file.ts";
import { runFlowFile, vacuousSteps } from "../support/fuzz.ts";
import {
  assertRequestedOutcome,
  capture,
  compareComments,
  expectOperation,
  modelOf,
} from "../support/oracle.ts";
import { stageReviewedProtectedEdit } from "../support/protected-edit.ts";

/** Read authored marker ownership independently of either comment projection. */
const markerOwners = (document: unknown, commentId: number) => {
  const owners = {
    start: [] as string[],
    end: [] as string[],
    reference: [] as string[],
  };
  const visit = (value: unknown, paragraphId: string | undefined): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child, paragraphId);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const owner =
      "type" in value &&
      value.type === "paragraph" &&
      "paraId" in value &&
      typeof value.paraId === "string"
        ? value.paraId
        : paragraphId;
    if (
      "id" in value &&
      value.id === commentId &&
      "type" in value &&
      (value.type === "commentRangeStart" ||
        value.type === "commentRangeEnd" ||
        value.type === "commentReference")
    ) {
      assert.ok(owner, "An authored comment marker must belong to a paragraph");
      switch (value.type) {
        case "commentRangeStart":
          owners.start.push(owner);
          break;
        case "commentRangeEnd":
          owners.end.push(owner);
          break;
        case "commentReference":
          owners.reference.push(owner);
          break;
      }
    }
    for (const child of Object.values(value)) visit(child, owner);
  };
  visit(document, undefined);
  return owners;
};

const cases = [
  {
    file: "issue-1842-comment-on-split-note-tail.json",
    story: { type: "footnote", noteId: 1 },
    applied: [2, 1],
  },
  {
    file: "issue-1842-comment-on-split-replacement-tail.json",
    story: { type: "main" },
    applied: [1, 1, 1],
  },
] as const;

for (const { file, story, applied } of cases)
  test(`${file} retires the accepted tail while retaining its pending live range`, async () => {
    const flow = parseFlowFile(
      JSON.parse(
        readFileSync(new URL(`../flows/${file}`, import.meta.url), "utf8"),
      ),
    );
    const run = await runFlowFile(
      { ...flow, steps: flow.steps.slice(0, -1) },
      { captureSaved: true },
    );
    assert.deepEqual(vacuousSteps(run), []);
    assert.deepEqual(
      run.effects.map((effect) => {
        assert.equal(effect.type, "batch");
        if (effect.type !== "batch") assert.fail("Expected an edit batch");
        return effect.applied;
      }),
      applied,
    );
    assert.ok(run.saved);
    const reviewer = await openReviewer(run.saved);
    const storyContent = () => {
      const document = reviewer.toDocument();
      return story.type === "main"
        ? document.package.document.content
        : document.package.footnotes?.find(({ id }) => id === story.noteId)
            ?.content;
    };
    const before = await capture(reviewer, "tracked-changes", { story });
    const [accepted] = before.comments;
    const [live] = before.liveComments;
    assert.ok(accepted?.blockId && live);
    assert.equal(before.comments.length, 1);
    assert.equal(before.liveComments.length, 1);
    assert.equal(accepted.anchor, "updated");
    assert.notEqual(live.blockId, accepted.blockId);
    assert.ok(live.anchor.endsWith(accepted.anchor));
    assert.deepEqual(markerOwners(storyContent(), accepted.id), {
      start: [live.blockId],
      end: [accepted.blockId],
      reference: [accepted.blockId],
    });
    const model = modelOf(before.rows, before.liveRows, before.joinSources);
    assert.equal(
      model.pendingJoins.size,
      0,
      "A tracked split is not a pending join",
    );

    const operation = {
      id: "delete-tail",
      type: "deleteBlock",
      blockId: accepted.blockId,
    } as const;
    stageReviewedProtectedEdit({ reviewer, story, operation, kind: "comment" });
    await assertRequestedOutcome(
      reviewer,
      before,
      { applied: [operation] },
      "review split tail deletion",
    );
    const after = await capture(reviewer, "tracked-changes", { story });
    assert.deepEqual(after.comments, []);
    assert.equal(after.liveComments.length, 1);
    assert.equal(after.liveComments.at(0)?.id, accepted.id);
    const owners = markerOwners(storyContent(), accepted.id);
    assert.equal(owners.start.length, 1);
    assert.equal(owners.end.length, 1);
    assert.equal(owners.reference.length, 1);
    expectOperation(model, operation);
    const issues = (comments: typeof after.comments) =>
      compareComments(
        model,
        before.comments,
        comments,
        after.rows,
        before.liveComments,
        "tracked-changes",
      );
    assert.deepEqual(issues([]), []);
    assert.match(issues([accepted]).join("\n"), /outlived/u);
  });

test("a differing live locator cannot excuse unrelated accepted comment loss", () => {
  const source = {
    id: "412DDC7F",
    text: "retained prefix",
    kind: "paragraph",
  } as const;
  const tail = { id: "483CD5EB", text: "updated", kind: "paragraph" } as const;
  const model = modelOf([source, tail]);
  expectOperation(model, { type: "deleteBlock", blockId: tail.id });
  const comment = {
    id: 1,
    text: "Keep the prefix",
    anchor: source.text,
    blockId: source.id,
  };
  const live = { ...comment, anchor: tail.text, blockId: tail.id };
  assert.equal(
    compareComments(model, [comment], [], [source], [live], "tracked-changes")
      .length,
    1,
  );
  assert.deepEqual(
    compareComments(
      model,
      [comment],
      [comment],
      [source],
      [live],
      "tracked-changes",
    ),
    [],
  );
});
