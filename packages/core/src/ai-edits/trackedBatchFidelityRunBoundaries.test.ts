/**
 * A tracked batch's Reject All proof reads where a run boundary falls as no
 * change: a comment added over words that end at a line break leaves, once
 * its markers are taken out, the paragraph the original was.
 */
import { expect, test } from "bun:test";

import type { Paragraph } from "../types/document";
import { sameOriginalBlocks } from "./tracked-batch-fidelity";

const original: Paragraph = {
  type: "paragraph",
  paraId: "1B0E5A11",
  content: [
    {
      type: "run",
      content: [
        { type: "text", text: "Delivery schedule" },
        { type: "break" },
        { type: "text", text: "Payment terms" },
      ],
    },
  ],
};

const commented: Paragraph = {
  type: "paragraph",
  paraId: "1B0E5A11",
  content: [
    { type: "run", content: [{ type: "text", text: "Delivery " }] },
    { type: "commentRangeStart", id: 0 },
    {
      type: "run",
      content: [{ type: "text", text: "schedule" }, { type: "break" }],
    },
    { type: "commentRangeEnd", id: 0 },
    { type: "commentReference", id: 0 },
    { type: "run", content: [{ type: "text", text: "Payment terms" }] },
  ],
};

test("a comment added across a line break leaves the original once taken out", () => {
  expect(sameOriginalBlocks([original], [commented], new Set([0]))).toBe(true);
});

test("runs of different properties stay apart", () => {
  const bold: Paragraph = structuredClone(commented);
  const run = bold.content[2];
  if (run?.type !== "run") throw new Error("the fixture has no run");
  run.formatting = { bold: true };
  expect(sameOriginalBlocks([original], [bold], new Set([0]))).toBe(false);
});
