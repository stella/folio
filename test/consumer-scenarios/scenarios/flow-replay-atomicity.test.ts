import assert from "node:assert/strict";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "../support/corpus.ts";
import { directNumberedDocument, openReviewer } from "../support/documents.ts";
import { type FlowFile, parseFlowFile } from "../support/flow-file.ts";
import { runFlowFile } from "../support/fuzz.ts";

test("seed 877798201 refuses insertion on a deleted anchor and rolls back only atomic batches", async () => {
  const pinned = readFlowDir(fileURLToPath(new URL("../flows/", import.meta.url))).find(
    ({ file }) => file === "suggested-insert-before-deleted-anchor.json",
  );
  assert.ok(pinned);
  for (const atomic of [false, true]) {
    const run = await runFlowFile({
      ...pinned.flow,
      steps: pinned.flow.steps.map((step) => Object.assign({}, step, { atomic })),
    });
    assert.deepEqual(run.effects, [{ type: "batch", applied: atomic ? 0 : 1 }]);
  }
});

test("pinned atomicity survives generator drift and controls whether a conflicting batch acts", async () => {
  const reviewer = await openReviewer(await directNumberedDocument());
  assert.equal(
    reviewer.getContent().find(({ id }) => id === "3CEFFC83")?.text,
    "Numbered clause one",
    "the exact replay target must exist before either batch",
  );
  for (const atomic of [false, true]) {
    const flow = {
      version: 1,
      kind: "collisions",
      generation: "targeted",
      fixture: "directNumbered",
      mode: "tracked-changes",
      seed: 118244301,
      steps: [
        {
          action: "core batch",
          seed: 1317974444,
          atomic,
          operations: [
            { type: "deleteBlock", blockId: "3CEFFC83" },
            { type: "splitBlock", blockId: "3CEFFC83", offset: 9 },
          ],
        },
      ],
    } as const satisfies FlowFile;
    const run = await runFlowFile(parseFlowFile(JSON.parse(JSON.stringify(flow))));
    assert.deepEqual(run.effects, [{ type: "batch", applied: atomic ? 0 : 1 }]);
    assert.equal(run.flow.steps.at(0)?.atomic, atomic);
  }
});
