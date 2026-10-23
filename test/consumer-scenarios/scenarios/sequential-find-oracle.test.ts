import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioDocxReviewer,
  hashFolioAIBlockText,
} from "@stll/folio-core/server";

import { openReviewer, packDocument } from "../support/documents.ts";
import {
  ENABLED_RELATIONS,
  relationCheckCount,
  startRelations,
} from "../support/metamorphic.ts";
import { mergedFindCollision } from "../support/operations.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const probeMergedFind = async (find: string) => {
  const fixture = await packDocument(
    fromMarkdown(
      `Before ${find} clause.\n\nAfter ${find} words.\n\nFinal survivor.`,
    ),
  );
  const reviewer = await openReviewer(fixture);
  const first = reviewer.getContent().at(0);
  assert.ok(first);
  const before = relationCheckCount("batchSequential");
  // This seed selects the first batch even when relation depth is sampled.
  const relations = await startRelations({
    fixture,
    reviewer,
    mode: "direct",
    seed: 14,
  });
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: first.id,
          separator: " ",
        },
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: first.id,
          find,
          replace: "revised",
          precondition: { blockTextHash: hashFolioAIBlockText(first.text) },
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(result.applied.length, 2);
  const bytes = new Uint8Array(await reviewer.toBuffer());
  await relations.afterStep(
    reviewer,
    { bytes, reopened: await openReviewer(bytes) },
    "pre-batch find anchor",
  );
  assert.equal(
    relationCheckCount("batchSequential"),
    before + 1,
    "the oracle must compare",
  );
  assert.equal(
    reviewer.getContent().at(0)?.text,
    `Before revised clause. After ${find} words.`,
  );
};

test(
  "sequential find replay keeps pre-batch anchors when merges create duplicate matches",
  { skip: !ENABLED_RELATIONS.has("batchSequential") },
  async () => {
    for (const find of ["needle", "👍🏽", "👩‍⚖️", "e\u0301"]) {
      await probeMergedFind(find);
    }
  },
);

test(
  "sequential find oracle rejects a batch replacement at the wrong pre-batch anchor",
  { skip: !ENABLED_RELATIONS.has("batchSequential") },
  async () => {
    const prototype = FolioDocxReviewer.prototype;
    const apply = prototype.applyDocumentOperations;
    // Defect affects real batches and the oracle's batch copies, while the
    // one-at-a-time replay still applies the requested pre-batch span.
    prototype.applyDocumentOperations = function (
      this: Reviewer,
      ...[batch, options]: Parameters<Reviewer["applyDocumentOperations"]>
    ) {
      return apply.call(
        this,
        {
          ...batch,
          operations: batch.operations.map((operation) =>
            batch.operations.length > 1 && operation.type === "replaceInBlock"
              ? Object.assign({}, operation, { find: "clause." })
              : operation,
          ),
        },
        options,
      );
    };
    try {
      await assert.rejects(
        probeMergedFind("needle"),
        /\[batchSequential\].*leaves otherwise/su,
      );
    } finally {
      prototype.applyDocumentOperations = apply;
    }
  },
);

test("duplicate-find collision generation names a unique pre-batch occurrence", () => {
  for (const find of ["needle", "Cláusula", "Clause"]) {
    const first = {
      id: "first",
      kind: "paragraph",
      text: `${find} begins here.`,
    } as const;
    const next = {
      id: "next",
      kind: "paragraph",
      text: `${find} follows there.`,
    } as const;
    const operations = mergedFindCollision([first, next]);
    assert.ok(operations);
    assert.equal(operations.at(0)?.type, "mergeBlockWithNext");
    assert.deepEqual(operations.at(1), {
      type: "replaceInBlock",
      blockId: first.id,
      find,
      replace: "revised",
    });
    assert.equal(first.text.indexOf(find, first.text.indexOf(find) + 1), -1);
    assert.notEqual(
      `${first.text} ${next.text}`.indexOf(find, first.text.length),
      -1,
    );
  }
  assert.equal(
    mergedFindCollision([
      { id: "first", kind: "paragraph", text: "Repeated Repeated words." },
      { id: "next", kind: "paragraph", text: "Repeated follows." },
    ]),
    null,
  );
});
