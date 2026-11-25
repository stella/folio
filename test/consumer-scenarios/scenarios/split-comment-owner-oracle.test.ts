import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { parseFlowFile } from "../support/flow-file.ts";
import { runFlowFile, vacuousSteps } from "../support/fuzz.ts";
import {
  capture,
  compareComments,
  expectOperation,
  modelOf,
} from "../support/oracle.ts";

const cases = [
  {
    file: "issue-1842-comment-on-split-note-tail.json",
    story: { type: "footnote", noteId: 1 },
    applied: [2, 1, 1],
    replaced: "the",
  },
  {
    file: "issue-1842-comment-on-split-replacement-tail.json",
    story: { type: "main" },
    applied: [1, 1, 1, 1],
    replaced: "Schedule",
  },
] as const;

const readFlow = (file: string) =>
  parseFlowFile(
    JSON.parse(
      readFileSync(new URL(`../flows/${file}`, import.meta.url), "utf8"),
    ),
  );

const appliedCounts = (run: Awaited<ReturnType<typeof runFlowFile>>) =>
  run.effects.map((effect) => {
    assert.equal(effect.type, "batch");
    if (effect.type !== "batch")
      assert.fail("Every saved step must apply a batch");
    return effect.applied;
  });

for (const { file, story, applied, replaced } of cases)
  test(`${file} keeps the replaced words' thread rejectable through the split and the tail deletion`, async () => {
    const flow = readFlow(file);
    let commentId: number | undefined;
    const observed = new Set<number>();
    const run = await runFlowFile(flow, {
      checks: [
        {
          name: "replaced comment ownership",
          check: async ({ flow: current, index }) => {
            if (index < flow.steps.length - 2) return;
            observed.add(index);
            const snapshot = await capture(
              current.reviewer,
              "tracked-changes",
              { story },
            );
            // The comment went with the words the replacement deleted: no
            // accepted owner holds it, and its live thread stays rejectable.
            assert.deepEqual(snapshot.comments, []);
            assert.equal(snapshot.liveComments.length, 1);
            const live = snapshot.liveComments.at(0);
            assert.ok(live);
            assert.equal(live.anchor, replaced);
            commentId ??= live.id;
            assert.equal(live.id, commentId);
          },
        },
      ],
    });
    assert.deepEqual(vacuousSteps(run), []);
    assert.deepEqual(appliedCounts(run), applied);
    assert.deepEqual(
      [...observed],
      [flow.steps.length - 2, flow.steps.length - 1],
    );
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
