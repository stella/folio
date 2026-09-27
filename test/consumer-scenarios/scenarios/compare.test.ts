/**
 * Compare round trips: a redline of before vs after is a document whose
 * accept-all is the after and whose reject-all is the before, which saves,
 * reopens and reads alike.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { compareDocxVersions, formatVersionDiffForLLM } from "@stll/folio-agents";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { generateRedlineDocx } from "@stll/folio-core/server";

import {
  FIXTURE_NAMES,
  FIXTURES,
  openReviewer,
  packDocument,
  toArrayBuffer,
} from "../support/documents.ts";
import { labelsOf } from "../support/editor.ts";
import { assertHealthy, assertReadersAgree } from "../support/invariants.ts";
import { type Block, coreBatch, randomOperation } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";
import { resolvedText } from "../support/review.ts";

/** The fixture with its own pending changes accepted: the text both sides compare. */
const settled = async (bytes: Uint8Array): Promise<Uint8Array> => {
  const reviewer = await openReviewer(bytes);
  reviewer.acceptAll();
  return new Uint8Array(await reviewer.toBuffer());
};

/** Direct edits a person makes between two versions. */
const edited = async (bytes: Uint8Array, seed: number): Promise<Uint8Array> => {
  const reviewer = await openReviewer(bytes);
  const random = createRandom(seed);
  const types = [
    "replaceInBlock",
    "insertAfterBlock",
    "replaceBlock",
    "deleteBlock",
    "insertTableRow",
  ];
  for (let step = 0; step < 4; step += 1) {
    const operation = randomOperation(reviewer.getContent() as Block[], "direct", random, types);
    if (operation) {
      reviewer.applyDocumentOperations(coreBatch([operation], "direct") as never);
    }
  }
  return new Uint8Array(await reviewer.toBuffer());
};

describe("compare round trips", () => {
  for (const [index, name] of FIXTURE_NAMES.entries()) {
    test(`${name}: redline(before, after) accepts to after and rejects to before`, async () => {
      const before = await settled(await FIXTURES[name]());
      const after = await edited(before, 1000 + index);

      const redline = await generateRedlineDocx(toArrayBuffer(before), toArrayBuffer(after), {
        author: "Comparison",
      });
      assert.deepEqual(redline.unprocessedStories, [], "the comparison skipped a story");
      const redlineBytes = new Uint8Array(redline.buffer);
      await assertReadersAgree(redlineBytes, `${name} redline`);
      await assertHealthy(await openReviewer(redlineBytes), `${name} redline saved`);

      assert.deepEqual(
        await resolvedText(redlineBytes, "accept"),
        await resolvedText(after, "accept"),
        `${name}: accepting the redline does not give the revised version`,
      );
      assert.deepEqual(
        await resolvedText(redlineBytes, "reject"),
        await resolvedText(before, "accept"),
        `${name}: rejecting the redline does not give the original version`,
      );

      const diff = await compareDocxVersions(toArrayBuffer(before), toArrayBuffer(after));
      assert.equal(typeof formatVersionDiffForLLM(diff), "string");
    });
  }

  test("accepting a redline keeps an inserted bullet a bullet", async () => {
    const before = await packDocument(fromMarkdown("Intro.\n\nOutro."));
    const after = await packDocument(fromMarkdown("Intro.\n\n- new bullet\n\nOutro."));
    const redline = await generateRedlineDocx(toArrayBuffer(before), toArrayBuffer(after));
    const reviewer = await openReviewer(new Uint8Array(redline.buffer));
    reviewer.acceptAll();
    assert.deepEqual(
      await labelsOf(new Uint8Array(await reviewer.toBuffer())),
      ["· Intro.", "• new bullet", "· Outro."],
      "the inserted list item lost its bullet",
    );
    assert.deepEqual(
      await resolvedText(new Uint8Array(redline.buffer), "reject"),
      await resolvedText(before, "accept"),
      "rejecting the redline does not give the original version",
    );
  });
});
