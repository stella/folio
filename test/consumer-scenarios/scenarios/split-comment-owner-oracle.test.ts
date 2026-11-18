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
    applied: [2, 1, 1],
  },
  {
    file: "issue-1842-comment-on-split-replacement-tail.json",
    story: { type: "main" },
    applied: [1, 1, 1, 1],
  },
] as const;

for (const { file, story, applied } of cases)
  test(`${file} retires the accepted tail while retaining its pending live range`, async () => {
    const flow = parseFlowFile(
      JSON.parse(
        readFileSync(new URL(`../flows/${file}`, import.meta.url), "utf8"),
      ),
    );
    let captured: Awaited<ReturnType<typeof capture>> | undefined;
    let commentId: number | undefined;
    const observed = new Set<number>();
    const run = await runFlowFile(flow, {
      checks: [
        {
          name: "independent split-range ownership",
          check: async ({ flow: current, index }) => {
            if (index < flow.steps.length - 2) return;
            observed.add(index);
            const snapshot = await capture(
              current.reviewer,
              "tracked-changes",
              { story },
            );
            const document = current.reviewer.toDocument();
            // Read the authored story content once; sections also project the
            // body paragraphs and must not count the same markers a second time.
            const content =
              story.type === "main"
                ? document.package.document.content
                : document.package.footnotes?.find(
                    ({ id }) => id === story.noteId,
                  )?.content;
            assert.ok(content);
            if (index === flow.steps.length - 2) {
              captured = snapshot;
              const accepted = snapshot.comments.at(0);
              const live = snapshot.liveComments.at(0);
              assert.ok(accepted && live);
              assert.equal(snapshot.comments.length, 1);
              assert.equal(snapshot.liveComments.length, 1);
              commentId = accepted.id;
              assert.equal(accepted.anchor, "updated");
              assert.notEqual(live.blockId, accepted.blockId);
              assert.ok(live.anchor.endsWith(accepted.anchor));
              assert.deepEqual(markerOwners(content, accepted.id), {
                start: [live.blockId],
                end: [accepted.blockId],
                reference: [accepted.blockId],
              });
              assert.equal(
                modelOf(snapshot.rows, snapshot.liveRows, snapshot.joinSources)
                  .pendingJoins.size,
                0,
                "A tracked split is not a pending paragraph join",
              );
              return;
            }
            assert.ok(captured && commentId !== undefined);
            assert.deepEqual(snapshot.comments, []);
            assert.equal(
              snapshot.liveComments.length,
              1,
              "Deleting the accepted coverage must keep the rejectable live thread",
            );
            assert.equal(snapshot.liveComments.at(0)?.id, commentId);
            const owners = markerOwners(content, commentId);
            assert.equal(owners.start.length, 1);
            assert.equal(owners.end.length, 1);
            assert.equal(owners.reference.length, 1);
            const accepted = captured.comments.at(0);
            assert.ok(accepted?.blockId);
            const model = modelOf(
              captured.rows,
              captured.liveRows,
              captured.joinSources,
            );
            expectOperation(model, {
              type: "deleteBlock",
              blockId: accepted.blockId,
            });
            assert.deepEqual(
              compareComments(
                model,
                captured.comments,
                [],
                snapshot.rows,
                captured.liveComments,
                "tracked-changes",
              ),
              [],
            );
            assert.match(
              compareComments(
                model,
                captured.comments,
                [accepted],
                snapshot.rows,
                captured.liveComments,
                "tracked-changes",
              ).join("\n"),
              /outlived/u,
              "Retaining the thread on its removed accepted owner is an orphan",
            );
          },
        },
      ],
    });
    assert.deepEqual(vacuousSteps(run), []);
    assert.deepEqual(
      run.effects.map((effect) => {
        assert.equal(effect.type, "batch");
        if (effect.type !== "batch")
          assert.fail("Every saved step must apply a batch");
        return effect.applied;
      }),
      applied,
    );
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
