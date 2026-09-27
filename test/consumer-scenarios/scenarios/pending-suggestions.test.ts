/** A host can persist proposals beside the DOCX and restore them after reopening it. */

import assert from "node:assert/strict";
import { test } from "node:test";

import { FOLIO_PENDING_SUGGESTION_VERSION } from "@stll/folio-core/ai-edits";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";

import { openReviewer, plainDocument } from "../support/documents.ts";

test("pending suggestions round-trip through the packed public API", async () => {
  const reviewer = await openReviewer(await plainDocument());
  const original = reviewer.getContent().map(({ text }) => text);
  const target = reviewer.getContent().find(({ text }) => text.includes("thirty days"));
  assert.ok(target);

  const applied = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "suggested",
    operations: [
      {
        id: "payment-proposal",
        type: "replaceInBlock",
        blockId: target.id,
        find: "thirty days",
        replace: "fourteen days",
      },
    ],
  });
  assert.equal(applied.applied.length, 1);
  const proposed = reviewer.getContent().map(({ text }) => text);
  assert.notDeepEqual(proposed, original);

  const records = reviewer.exportPendingSuggestions();
  assert.equal(records.length, 1);
  assert.equal(records[0]?.version, FOLIO_PENDING_SUGGESTION_VERSION);
  const stored: unknown = JSON.parse(JSON.stringify(records));
  assert.ok(Array.isArray(stored));

  const reopened = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  assert.deepEqual(
    reopened.getContent().map(({ text }) => text),
    original,
  );
  assert.deepEqual(reopened.loadPendingSuggestions(stored), [
    { status: "restaged", suggestionId: "payment-proposal" },
  ]);
  assert.deepEqual(
    reopened.getContent().map(({ text }) => text),
    proposed,
  );
  assert.equal(reopened.acceptAll(), 0);
  assert.deepEqual(
    reopened.getContent().map(({ text }) => text),
    proposed,
  );

  const stillPending = await openReviewer(new Uint8Array(await reopened.toBuffer()));
  assert.deepEqual(
    stillPending.getContent().map(({ text }) => text),
    original,
  );
  assert.equal(reopened.acceptSuggestion("payment-proposal"), true);
  const accepted = await openReviewer(new Uint8Array(await reopened.toBuffer()));
  assert.deepEqual(
    accepted.getContent().map(({ text }) => text),
    proposed,
  );

  const unsupported = records.map((record) => Object.assign({}, record, { version: 999 }));
  const clean = await openReviewer(await plainDocument());
  assert.deepEqual(clean.loadPendingSuggestions(unsupported), [
    { status: "stale", suggestionId: "payment-proposal", reason: "unsupportedVersion" },
  ]);
  assert.deepEqual(
    clean.getContent().map(({ text }) => text),
    original,
  );
});
