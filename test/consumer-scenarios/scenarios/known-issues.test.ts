/**
 * The open issues and unfiled findings, each as the smallest public-API
 * scenario that shows it. Every one runs as an expected failure: when a fix
 * lands, its scenario passes, the expected failure fails, and the marker comes
 * off so the scenario guards the fix.
 */

import assert from "node:assert/strict";
import { describe } from "node:test";

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/server";

import { openReviewer, plainDocument, unusedNumberingDocument } from "../support/documents.ts";
import { saveAndReopen } from "../support/invariants.ts";
import { runFlow } from "../support/fuzz.ts";
import { expectedFailure, KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";
import { MODES } from "../support/operations.ts";

const MISSING_NUMBERING = /Numbering definition \d+ is missing/u;

describe("#1103: operations that name a numbering instance the package does not define", () => {
  for (const mode of MODES) {
    expectedFailure(
      1103,
      `insertAfterBlock with an undefined numId (${mode}) is refused or saves`,
      MISSING_NUMBERING,
      async () => {
        const reviewer = await openReviewer(await unusedNumberingDocument());
        const anchor = reviewer
          .getContent()
          .find((block) => block.text === "Signed in two copies.");
        assert.ok(anchor);
        reviewer.applyDocumentOperations({
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "1",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "An inserted clause.",
              numbering: { numId: 1, level: 0 },
            },
          ],
        });
        if (mode === "suggested") {
          // A suggestion reaches the package once accepted.
          reviewer.acceptAll();
        }
        await saveAndReopen(reviewer, `#1103 ${mode}`);
      },
    );
  }

  expectedFailure(
    1103,
    "setBlockParagraphProperties with an undefined numId is refused or saves",
    MISSING_NUMBERING,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const target = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
      assert.ok(target);
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "1",
            type: "setBlockParagraphProperties",
            blockId: target.id,
            properties: { numbering: { numId: 3, level: 0 } },
          },
        ],
      });
      await saveAndReopen(reviewer, "#1103 setBlockParagraphProperties");
    },
  );

  expectedFailure(
    1103,
    "suggest_changes with an undefined numId answers with an issue, and the document saves",
    MISSING_NUMBERING,
    async () => {
      const reviewer = await openReviewer(await plainDocument());
      const anchor = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
      assert.ok(anchor);
      executeFolioToolCallUntyped(
        "suggest_changes",
        {
          operations: [
            {
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "An inserted clause.",
              numbering: { numId: 1, level: 0 },
            },
          ],
        },
        createReviewerBridge(reviewer, { mode: "tracked-changes" }),
        {},
      );
      await saveAndReopen(reviewer, "#1103 suggest_changes");
    },
  );
});

describe("findings not yet filed", () => {
  for (const { seed, steps, finding } of KNOWN_FAILING_FLOWS) {
    expectedFailure(
      finding,
      `the fuzz flow with seed ${seed} (${steps} steps) saves what the reviewer shows`,
      /reopened package shows something else/u,
      () => runFlow(seed, steps),
    );
  }
});
