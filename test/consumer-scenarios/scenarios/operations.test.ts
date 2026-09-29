/**
 * `FolioDocxReviewer.applyDocumentOperations` over every fixture, in every
 * mode, one operation type at a time: each batch applies or refuses with an
 * issue, an applied one did what it asked (support/oracle.ts), the result
 * saves and reopens, every reader agrees, and resolving ordinary tracked
 * changes lands where it should (reject-all: the document before; accept-all:
 * the document a reader saw). Suggested edits stay in the host store.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";

import {
  directNumberedDocument,
  FIXTURE_NAMES,
  FIXTURES,
  listDocument,
  openReviewer,
  plainDocument,
  unusedNumberingDocument,
} from "../support/documents.ts";
import { assertHealthy, saveAndReopen, visibleState } from "../support/invariants.ts";
import { reportScenarioFailure, shellQuote } from "../support/failure-fingerprints.ts";
import { assertRequestedOutcome, capture } from "../support/oracle.ts";
import { type Block, coreBatch, GENERATORS, MODES, supports } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";
import {
  expectedFailure,
  FINDING_SYMPTOMS,
  KNOWN_FAILING_OPERATION_RUNS,
} from "../support/known-issues.ts";
import { resolvedText, settledText } from "../support/review.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const blocksOf = (reviewer: Reviewer): Block[] => reviewer.getContent() as Block[];

const matrixSeed = (fallback: number): number => {
  const raw = process.env["FOLIO_OPERATION_MATRIX_SEED"];
  if (raw === undefined) return fallback;
  const seed = Number(raw);
  if (!Number.isSafeInteger(seed)) throw new Error(`Invalid operation matrix seed: ${raw}`);
  return seed;
};

const testPattern = (title: string): string => `^${title.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")}`;

describe("applyDocumentOperations", () => {
  for (const name of FIXTURE_NAMES) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_OPERATION_RUNS.find(
        (run) => run.fixture === name && run.mode === mode,
      );
      const title = `${name} / ${mode}: every operation type applies or refuses, and the result saves`;
      const seed = matrixSeed(name.length * 31 + mode.length);
      const repro = `FOLIO_OPERATION_MATRIX_SEED=${seed} bun scripts/consumer-scenarios.ts --only ${shellQuote(testPattern(title))}`;
      const register = (body: () => Promise<void>) => {
        const run = async () => {
          try {
            await body();
          } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (known && FINDING_SYMPTOMS[known.finding].test(message)) throw error;
            reportScenarioFailure({
              test: `operation matrix ${name} / ${mode}`,
              seed,
              repro,
              failure: error,
            });
          }
        };
        if (known) {
          return expectedFailure(known.finding, title, FINDING_SYMPTOMS[known.finding], run);
        }
        return test(title, run);
      };
      register(async () => {
        const before = await FIXTURES[name]();
        const reviewer = await openReviewer(before);
        // `"suggested"` edits stay out of the saved package until accepted.
        const persisted = mode === "suggested" ? { persisted: visibleState(reviewer) } : {};
        const random = createRandom(seed);
        const log: string[] = [];
        for (const type of Object.keys(GENERATORS)) {
          if (!supports(type, mode)) continue;
          const operation = GENERATORS[type]?.(blocksOf(reviewer), random);
          if (!operation) continue;
          const pre = await capture(reviewer, mode);
          // These are separate batches; keep their proposal ids distinct across the run.
          const result = reviewer.applyDocumentOperations(
            coreBatch([{ ...operation, id: `op-${type}` }], mode) as never,
          );
          const outcome =
            result.applied.length > 0
              ? "applied"
              : `refused ${result.issues.map((issue) => issue.code).join(",")}`;
          log.push(`${type}: ${outcome}`);
          // What was asked, after a save and a reopen (accepted, when tracked).
          await assertRequestedOutcome(
            reviewer,
            pre,
            { applied: result.applied.length > 0 ? [operation] : [] },
            `${name} / ${mode}: ${JSON.stringify(operation)}`,
          );
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
          assert.ok(reviewer.exportPendingSuggestions().length > 0);
        }
        const { bytes: after } = await assertHealthy(
          reviewer,
          `${name} / ${mode} final`,
          persisted,
        );
        // For ordinary tracked changes, rejecting gives the document back,
        // block for block. Accepting lands on the words a reader saw; block
        // boundaries are left out there, as
        // a reader shows a pending join or split as the blocks it has now,
        // and so is whitespace: a merge's separator belongs to neither block
        // until the join is accepted.
        const words = (blocks: string[]) =>
          blocks
            .map((block) => block.replace(/^\w+: /u, ""))
            .join("")
            .replace(/\s+/gu, "");
        if (mode === "tracked-changes") {
          assert.deepEqual(
            await resolvedText(after, "reject"),
            await resolvedText(before, "reject"),
            `${name} / ${mode}: rejecting every change does not give the document back`,
          );
        }
        if (mode !== "suggested") {
          assert.equal(
            words(await resolvedText(after, "accept")),
            words(settledText(await openReviewer(after))),
            `${name} / ${mode}: accepting every change does not give what readers showed`,
          );
        }
      });
    }
  }
});

describe("a batch that splits a block and deletes it", () => {
  test("refuses the deletion, whose block the split already claims, and loses no word", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const text = "The Buyer pays each invoice within thirty days.";
    const target = reviewer.getContent().find((block) => block.text === text);
    assert.ok(target);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "split", type: "splitBlock", blockId: target.id, offset: 20 },
        { id: "delete", type: "deleteBlock", blockId: target.id },
      ],
    });
    assert.deepEqual(
      result.applied.map(({ id }) => id),
      ["split"],
    );
    assert.deepEqual(
      result.issues.map(({ operationId, code }) => ({ operationId, code })),
      [{ operationId: "delete", code: "overlappingOperation" }],
    );
    reviewer.acceptAll();
    const texts = reviewer.getContent().map((block) => block.text);
    assert.ok(
      texts.includes(text.slice(0, 20)) && texts.includes(text.slice(20)),
      `the split block's two halves: ${JSON.stringify(texts)}`,
    );
  });
});

describe("resolving tracked edits that build on pending ones", () => {
  const SUPPLIER = "The Supplier delivers the goods on time and in good order.";
  const tracked = async () => {
    const reviewer = await openReviewer(await plainDocument());
    const apply = (operation: Record<string, unknown>) =>
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "1", ...operation }],
      } as never);
    const block = (prefix: string) => {
      const found = reviewer.getContent().find(({ text }) => text.startsWith(prefix));
      assert.ok(found, `no block starts with "${prefix}"`);
      return found;
    };
    return { reviewer, apply, block };
  };
  const originalText = async () =>
    (await openReviewer(await plainDocument())).getContent().map(({ text }) => text);

  test("rejectAll undoes a merge of a split's second half into an inserted paragraph", async () => {
    const { reviewer, apply, block } = await tracked();
    apply({ type: "insertAfterBlock", blockId: block("Signed").id, text: "Inserted clause." });
    apply({ type: "splitBlock", blockId: block("Signed").id, offset: "Signed in ".length });
    apply({ type: "mergeBlockWithNext", blockId: block("two copies").id, separator: " " });
    reviewer.rejectAll();
    assert.deepEqual(
      reviewer.getContent().map(({ text }) => text),
      await originalText(),
    );
  });

  test("rejecting a split with a table inserted between its halves joins them again", async () => {
    const { reviewer, apply, block } = await tracked();
    const target = block("The Supplier");
    apply({ type: "splitBlock", blockId: target.id, offset: SUPPLIER.indexOf("good order") });
    apply({ type: "insertTable", blockId: target.id, rows: [["Term", "Value"]] });
    reviewer.rejectAll();
    assert.deepEqual(
      reviewer.getContent().map(({ text }) => text),
      await originalText(),
    );
  });

  test("a comment's anchored text reads the same before and after a save when its text is replaced", async () => {
    const { reviewer, apply, block } = await tracked();
    const target = block("The Supplier");
    const start = SUPPLIER.indexOf("good order");
    const range = createFolioAITextRangeHandle({
      blockId: target.id,
      text: SUPPLIER,
      startOffset: start,
      endOffset: start + "good".length,
    });
    assert.ok(range);
    apply({ type: "commentOnRange", range, comment: { text: "Which standard?" } });
    apply({ type: "replaceBlock", blockId: target.id, text: "The Supplier delivers promptly." });
    const before = reviewer.getComments().map((comment) => comment.anchoredText);
    // The removed word, and the new one that takes its place.
    assert.deepEqual(before, ["goodpromptly"]);
    const { reopened } = await saveAndReopen(reviewer, "comment anchor");
    assert.deepEqual(
      reopened.getComments().map((comment) => comment.anchoredText),
      before,
      "the anchored text changed across the save",
    );
  });
});

describe("list labels after an operation", () => {
  test("an item inserted into a list reads its own number, and the items after it renumber, before a save", async () => {
    const reviewer = await openReviewer(await listDocument());
    const anchor = reviewer.getContent().find((block) => block.text === "Deposit on signature");
    assert.ok(anchor);
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        { id: "1", type: "insertAfterBlock", blockId: anchor.id, text: "Interim payment" },
      ],
    });
    const live = reviewer
      .getContent()
      .filter((block) => block.listReference?.numId === anchor.listReference?.numId)
      .map((block) => `${block.displayLabel} ${block.text}`);
    assert.deepEqual(
      live,
      [
        "1. Deposit on signature",
        "2. Interim payment",
        "3. Balance on delivery",
        "4. Retention after inspection",
      ],
      "labels before the save are stale",
    );
    await assertHealthy(reviewer, "insert into a list");
  });

  test("a paragraph numbered at a level its instance does not define reads alike everywhere", async () => {
    const reviewer = await openReviewer(await directNumberedDocument());
    const anchor = reviewer.getContent().find(({ text }) => text === "Unnumbered body text.");
    assert.ok(anchor);
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "insertAfterBlock",
          blockId: anchor.id,
          text: "Level eight.",
          numbering: { numId: 7, level: 8 },
        },
      ],
    });
    // No marker is painted for it: prose that keeps its level.
    const unmarked = reviewer.getContent().find(({ text }) => text === "Level eight.");
    assert.equal(unmarked?.kind, "paragraph");
    assert.equal(unmarked?.displayLabel, undefined);
    assert.equal(unmarked?.listLevel, 8);
    await assertHealthy(reviewer, "undefined level");
  });
});

describe("an operation naming a paragraph style the package does not define", () => {
  for (const mode of MODES) {
    for (const styleId of ["NoSuchStyle", "TableGrid"]) {
      test(`insertAfterBlock with styleId ${styleId} is refused, not applied (${mode})`, async () => {
        const reviewer = await openReviewer(await plainDocument());
        const anchor = reviewer
          .getContent()
          .find((block) => block.text === "Signed in two copies.");
        assert.ok(anchor);
        const before = visibleState(reviewer);
        const result = reviewer.applyDocumentOperations({
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "1",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "An inserted clause.",
              styleId,
            },
          ],
        });
        assert.deepEqual(result.applied, []);
        assert.deepEqual(
          result.skipped.map(({ reason }) => reason),
          ["missingStyle"],
        );
        assert.deepEqual(visibleState(reviewer), before);
        await assertHealthy(reviewer, `undefined style ${styleId} (${mode})`);
      });
    }
  }

  test("setBlockParagraphProperties with an undefined styleId is refused, not applied", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const target = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
    assert.ok(target);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "1",
          type: "setBlockParagraphProperties",
          blockId: target.id,
          properties: { styleId: "NoSuchStyle" },
        },
      ],
    });
    assert.deepEqual(result.applied, []);
    assert.deepEqual(
      result.skipped.map(({ reason }) => reason),
      ["missingStyle"],
    );
    await assertHealthy(reviewer, "setBlockParagraphProperties undefined styleId");
  });

  test("control: a defined paragraph style applies and reopens as a heading", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const target = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
    assert.ok(target);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "1",
          type: "setBlockParagraphProperties",
          blockId: target.id,
          properties: { styleId: "Heading2" },
        },
      ],
    });
    assert.deepEqual(
      result.applied.map(({ id }) => id),
      ["1"],
    );
    const { reopened } = await assertHealthy(reviewer, "defined paragraph style");
    const heading = reopened.getContent().find((block) => block.text === "Signed in two copies.");
    assert.equal(heading?.kind, "heading");
    assert.equal(heading?.headingLevel, 2);
  });
});

describe("an operation naming a numbering instance the package does not define (#1103)", () => {
  for (const mode of MODES) {
    test(`insertAfterBlock with an undefined numId is refused, not saved broken (${mode})`, async () => {
      const reviewer = await openReviewer(await unusedNumberingDocument());
      const anchor = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
      assert.ok(anchor);
      const result = reviewer.applyDocumentOperations({
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
      assert.deepEqual(result.applied, []);
      assert.deepEqual(
        result.skipped.map(({ reason }) => reason),
        ["missingNumbering"],
      );
      if (mode === "suggested") {
        // Nothing was applied, so there is nothing to accept; the package
        // must still save and reopen unchanged.
        reviewer.acceptAll();
      }
      await assertHealthy(reviewer, `undefined numId (${mode})`);
    });
  }

  test("setBlockParagraphProperties with an undefined numId is refused, not saved broken", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const target = reviewer.getContent().find((block) => block.text === "Signed in two copies.");
    assert.ok(target);
    const result = reviewer.applyDocumentOperations({
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
    assert.deepEqual(result.applied, []);
    assert.deepEqual(
      result.skipped.map(({ reason }) => reason),
      ["missingNumbering"],
    );
    await assertHealthy(reviewer, "setBlockParagraphProperties undefined numId");
  });
});

describe("a batch that merges a block into one it deletes", () => {
  // Direct and tracked-then-accepted leave the same paragraphs: the merge
  // joins across the deleted block, or, where the deletions run to the end
  // and nothing is left to join, directly it is refused and tracked it joins
  // no separator onto the end.
  const outcomes = async (mergedPrefix: string, deletedPrefixes: readonly string[]) => {
    const texts: string[][] = [];
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await openReviewer(await plainDocument());
      const idOf = (prefix: string) => {
        const block = reviewer.getContent().find((candidate) => candidate.text.startsWith(prefix));
        assert.ok(block, prefix);
        return block.id;
      };
      reviewer.applyDocumentOperations(
        coreBatch(
          [
            ...deletedPrefixes.map((prefix) => ({ type: "deleteBlock", blockId: idOf(prefix) })),
            { type: "mergeBlockWithNext", blockId: idOf(mergedPrefix), separator: " " },
          ],
          mode,
        ) as never,
      );
      texts.push(await resolvedText(new Uint8Array(await reviewer.toBuffer()), "accept"));
    }
    return texts;
  };

  test("joins across it, directly and tracked alike", async () => {
    const [direct, tracked] = await outcomes("This agreement", ["The Supplier"]);
    assert.deepEqual(tracked, direct);
  });

  test("with the deletions running to the end, leaves the merged block as it was", async () => {
    const [direct, tracked] = await outcomes("The Buyer", ["Signed"]);
    assert.deepEqual(tracked, direct);
    assert.ok(
      direct?.some((text) => text.endsWith("The Buyer pays each invoice within thirty days.")),
    );
  });
});
