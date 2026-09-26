/**
 * `suggest_changes` with a `numbering.numId` the document does not define
 * (issue #1103): the call succeeds, the operation is reported skipped with a
 * reason the model can act on, and the reviewer still saves.
 */

import { describe, expect, test } from "bun:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { createDocx, ensureParaIds, FolioDocxReviewer } from "@stll/folio-core/server";
import type { NumberingDefinitions } from "@stll/folio-core/types/document";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCallUntyped } from "./execute";

const ANCHOR = "1.2. Second clause.";

const UNUSED_901: NumberingDefinitions = {
  abstractNums: [
    {
      abstractNumId: 901,
      multiLevelType: "multilevel",
      levels: [{ ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1.", suffix: "tab" }],
    },
  ],
  nums: [{ numId: 901, abstractNumId: 901 }],
};

const openReviewer = async (numbering?: NumberingDefinitions) => {
  const model = fromMarkdown(`# Agreement\n\n1.1. First clause.\n\n${ANCHOR}\n\nTail.`);
  if (numbering) {
    model.package.numbering = numbering;
  }
  const { docx } = await ensureParaIds(await createDocx(model));
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, { author: "AI" });
  const anchor = reviewer.getContent().find((block) => block.text === ANCHOR)?.id;
  if (anchor === undefined) {
    throw new Error("fixture must expose its anchor block");
  }
  return { reviewer, anchor };
};

const texts = (reviewer: FolioDocxReviewer) => reviewer.getContent().map((block) => block.text);

describe("suggest_changes with an undefined numbering instance", () => {
  test.each([
    ["without a numbering part", undefined],
    ["with other instances only", UNUSED_901],
  ] as const)("is skipped, not applied, and the reviewer saves (%s)", async (_name, numbering) => {
    const { reviewer, anchor } = await openReviewer(numbering);
    const before = texts(reviewer);
    const out = executeFolioToolCallUntyped(
      "suggest_changes",
      {
        operations: [
          {
            type: "insertAfterBlock",
            blockId: anchor,
            text: "1.3. Inserted clause.",
            numbering: { numId: 1, level: 0 },
          },
        ],
      },
      createReviewerBridge(reviewer, { mode: "tracked-changes" }),
      {},
    );
    if (!out.ok) {
      throw new Error(out.error);
    }
    const result = out.result as {
      applied: unknown[];
      skipped: { id: string; reason: string }[];
      issues: { code: string; recovery: string; message?: string }[];
    };
    expect(result.applied).toEqual([]);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]?.reason).toContain("numbering instance this document does not define");
    expect(result.skipped[0]?.reason).toContain("numId 1");
    expect(result.issues).toEqual([
      expect.objectContaining({ code: "missingNumbering", recovery: "refreshDocument" }),
    ]);
    expect(texts(reviewer)).toEqual(before);
    await reviewer.toBuffer();
  });

  test("control: a defined instance applies and saves", async () => {
    const { reviewer, anchor } = await openReviewer(UNUSED_901);
    const out = executeFolioToolCallUntyped(
      "suggest_changes",
      {
        operations: [
          {
            type: "insertAfterBlock",
            blockId: anchor,
            text: "1.3. Inserted clause.",
            numbering: { numId: 901, level: 0 },
          },
        ],
      },
      createReviewerBridge(reviewer, { mode: "tracked-changes" }),
      {},
    );
    expect(out.ok).toBe(true);
    expect(out.ok && (out.result as { applied: unknown[] }).applied).toHaveLength(1);
    expect(texts(reviewer)).toContain("1.3. Inserted clause.");
    await reviewer.toBuffer();
  });
});
