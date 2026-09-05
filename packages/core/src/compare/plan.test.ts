/**
 * Move detection, at the planning layer.
 *
 * The benchmark corpus relocates paragraphs verbatim, so it proves the exact
 * path and says nothing about the threshold. A relocated paragraph is usually
 * edited on the way, and where the line sits — what still counts as the same
 * paragraph somewhere else, and what is a deletion plus an unrelated
 * insertion — is a judgement the plan makes and this file pins.
 */

import { describe, expect, test } from "bun:test";

import type { FolioAIBlock, FolioAIEditSnapshot } from "../ai-edits/types";
import { planStoryCompare } from "./plan";

const MAIN_STORY = { type: "main" } as const;

const block = (id: string, text: string): FolioAIBlock => ({ id, kind: "paragraph", text });

const snapshotOf = (blocks: readonly FolioAIBlock[]): FolioAIEditSnapshot => ({
  blocks: [...blocks],
  anchors: Object.fromEntries(
    blocks.map((entry, index) => [
      entry.id,
      {
        id: entry.id,
        from: index * 2,
        to: index * 2 + 1,
        text: entry.text,
        normalizedText: entry.text,
        textHash: entry.text,
        hashOccurrenceCount: 1,
      },
    ]),
  ),
});

const planOf = (base: readonly FolioAIBlock[], target: readonly FolioAIBlock[]) => {
  const plan = planStoryCompare({
    story: MAIN_STORY,
    baseSnapshot: snapshotOf(base),
    targetBlocks: target,
    maxOperations: 1000,
  });
  if (plan === null) {
    throw new Error("The plan exceeded its operation budget.");
  }
  return plan;
};

const RELOCATED =
  "The Supplier shall deliver the Goods to the named place within thirty days of the order.";

/** The same clause with one word changed: still the clause that moved. */
const RELOCATED_EDITED =
  "The Supplier shall deliver the Products to the named place within thirty days of the order.";

/** A clause that shares only its opening: a different obligation entirely. */
const UNRELATED =
  "The Supplier shall not be liable for any indirect loss however it arises in contract.";

describe("move detection", () => {
  test("a relocated paragraph is a move even when a word changed on the way", () => {
    const { changes, operations } = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", RELOCATED_EDITED)],
    );

    expect(changes.map(({ kind }) => kind)).toEqual(["move"]);
    // Both halves of the pair carry the same link, so the applier writes
    // `w:moveFrom` and `w:moveTo` rather than two unrelated revisions.
    const moveIds = operations.flatMap((operation) =>
      "moveId" in operation && operation.moveId !== undefined ? [operation.moveId] : [],
    );
    expect(moveIds).toHaveLength(2);
    expect(new Set(moveIds).size).toBe(1);
  });

  test("a paragraph that only shares its opening is a deletion and an insertion", () => {
    const { changes, operations } = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", UNRELATED)],
    );

    expect(changes.map(({ kind }) => kind).toSorted()).toEqual(["delete", "insert"]);
    expect(operations.some((operation) => "moveId" in operation)).toBe(false);
  });

  test("a short relocated line does not pair, so boilerplate does not read as a move", () => {
    const { changes } = planOf(
      [block("a", "Schedule 1"), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", "Schedule 1")],
    );

    expect(changes.map(({ kind }) => kind).toSorted()).toEqual(["delete", "insert"]);
  });
});
