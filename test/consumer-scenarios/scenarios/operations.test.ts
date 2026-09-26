/**
 * `FolioDocxReviewer.applyDocumentOperations` over every fixture, in every
 * mode, one operation type at a time: each batch applies or refuses with an
 * issue, the result saves and reopens, every reader agrees, and resolving the
 * changes lands where it should (reject-all: the document before; accept-all:
 * the document a reader saw).
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { FIXTURE_NAMES, FIXTURES, openReviewer } from "../support/documents.ts";
import { assertHealthy, visibleState } from "../support/invariants.ts";
import { type Block, coreBatch, GENERATORS, MODES, supports } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";
import { expectedFailure, KNOWN_FAILING_OPERATION_RUNS } from "../support/known-issues.ts";
import { resolvedText, settledText } from "../support/review.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const blocksOf = (reviewer: Reviewer): Block[] => reviewer.getContent() as Block[];

describe("applyDocumentOperations", () => {
  for (const name of FIXTURE_NAMES) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_OPERATION_RUNS.find(
        (run) => run.fixture === name && run.mode === mode,
      );
      const title = `${name} / ${mode}: every operation type applies or refuses, and the result saves`;
      const register = (body: () => Promise<void>) =>
        known
          ? expectedFailure(known.finding, title, /out of range|nodeSize/u, body)
          : test(title, body);
      register(async () => {
        const before = await FIXTURES[name]();
        const reviewer = await openReviewer(before);
        // `"suggested"` edits stay out of the saved package until accepted.
        const persisted = mode === "suggested" ? { persisted: visibleState(reviewer) } : {};
        const random = createRandom(name.length * 31 + mode.length);
        const log: string[] = [];
        for (const type of Object.keys(GENERATORS)) {
          if (!supports(type, mode)) continue;
          const operation = GENERATORS[type]?.(blocksOf(reviewer), random);
          if (!operation) continue;
          const result = reviewer.applyDocumentOperations(coreBatch([operation], mode) as never);
          const outcome =
            result.applied.length > 0
              ? "applied"
              : `refused ${result.issues.map((issue) => issue.code).join(",")}`;
          log.push(`${type}: ${outcome}`);
          assert.ok(
            result.applied.length + result.skipped.length === 1,
            `${type} neither applied nor refused: ${JSON.stringify(result)}`,
          );
          if (result.skipped.length > 0) {
            assert.ok(result.issues.length > 0, `${type} was skipped without an issue`);
          }
          await assertHealthy(reviewer, `${name} / ${mode} after ${log.join(" → ")}`, persisted);
        }
        assert.ok(
          log.some((entry) => entry.endsWith("applied")),
          `nothing applied: ${log.join("; ")}`,
        );

        if (mode === "suggested") {
          // Accepted, the suggestions are ordinary content and save as such.
          reviewer.acceptAll();
        }
        const { bytes: after } = await assertHealthy(reviewer, `${name} / ${mode} final`);
        // Resolving lands on the same words. Block boundaries are left out:
        // a reader shows a pending join or split as the blocks it has now,
        // and rejecting a split that has a table inserted after its first
        // half leaves it split (REJECT_SPLIT_AROUND_INSERTED_TABLE).
        // Whitespace is left out too: a merge's separator belongs to neither
        // block until the join is accepted.
        const words = (blocks: string[]) =>
          blocks
            .map((block) => block.replace(/^\w+: /u, ""))
            .join("")
            .replace(/\s+/gu, "");
        if (mode === "tracked-changes") {
          assert.equal(
            words(await resolvedText(after, "reject")),
            words(await resolvedText(before, "reject")),
            `${name} / ${mode}: rejecting every change does not give the document back`,
          );
        }
        assert.equal(
          words(await resolvedText(after, "accept")),
          words(settledText(await openReviewer(after))),
          `${name} / ${mode}: accepting every change does not give what readers showed`,
        );
      });
    }
  }
});
