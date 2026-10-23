import assert from "node:assert/strict";
import { test } from "node:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";

import {
  adjacentSectionRevisionDocument,
  openReviewer,
  randomMultiSectionDocument,
  storiesDocument,
} from "../support/documents.ts";
import { saveAndReopen } from "../support/invariants.ts";
import { createRandom } from "../support/random.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const sectionPropertiesOf = (reviewer: Reviewer) => {
  const document = reviewer.toDocument().package.document;
  return [
    ...document.content.flatMap((block) =>
      block.type === "paragraph" && block.sectionProperties !== undefined
        ? [{ paraId: block.paraId, properties: block.sectionProperties }]
        : [],
    ),
    { paraId: "body", properties: document.finalSectionProperties },
  ];
};

test("rejecting an added adjacent boundary keeps both sections after save and reopen", async () => {
  const reviewer = await openReviewer(await adjacentSectionRevisionDocument());

  reviewer.rejectAll();
  const { reopened } = await saveAndReopen(
    reviewer,
    "rejected adjacent section boundary",
  );

  assert.deepEqual(
    sectionPropertiesOf(reopened).map(
      ({ properties }) => properties?.pageWidth,
    ),
    [10_000, 12_000, 14_000],
  );
});

test("seed 1346488347: rejecting a deleted section boundary survives save and reopen", async () => {
  const reviewer = await openReviewer(await storiesDocument());
  const original = sectionPropertiesOf(reviewer);
  const boundary = reviewer
    .getContent()
    .find(({ text }) => text === "This clause ends the first section.");
  assert.ok(boundary);

  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "delete-boundary", type: "deleteBlock", blockId: boundary.id },
        {
          id: "insert-after-boundary",
          type: "insertAfterBlock",
          blockId: boundary.id,
          text: "Inserted section text.",
          alignment: "right",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(result.applied.length, 2);
  const pending = await saveAndReopen(
    reviewer,
    "pending section boundary deletion",
  );
  reviewer.rejectAll();
  pending.reopened.rejectAll();
  const { reopened } = await saveAndReopen(
    pending.reopened,
    "rejected section boundary deletion",
  );
  assert.deepEqual(sectionPropertiesOf(reviewer), original);
  assert.deepEqual(sectionPropertiesOf(reopened), original);
});

test("random tracked paragraph-range deletions restore every section after a real round trip", async () => {
  const random = createRandom(1_346_488_332);
  for (let run = 0; run < 24; run += 1) {
    const sectionCount = 3 + random.int(4);
    const reviewer = await openReviewer(
      await randomMultiSectionDocument(random.state(), sectionCount),
    );
    const original = sectionPropertiesOf(reviewer);
    const paragraphs = reviewer
      .getContent()
      .filter(({ table }) => table === undefined);
    const start = random.int(paragraphs.length);
    const length = 1 + random.int(Math.min(3, paragraphs.length - start));
    const selected = paragraphs.slice(start, start + length);
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: selected.map(({ id: blockId }, index) => ({
          id: `delete-${index}`,
          type: "deleteBlock" as const,
          blockId,
        })),
      },
      { undefinedReferences: "refuse" },
    );
    assert.equal(result.applied.length, selected.length);
    const { reopened: pending } = await saveAndReopen(
      reviewer,
      `pending section property run ${run}`,
    );
    pending.rejectAll();
    const { reopened } = await saveAndReopen(
      pending,
      `section property run ${run}`,
    );
    assert.deepEqual(sectionPropertiesOf(reopened), original);
  }
});
