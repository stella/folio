/**
 * `suggest_changes` with a `styleId` the document does not define as a
 * paragraph style: the call succeeds, the operation is reported skipped with
 * a reason the model can act on, and the reviewer still saves.
 */

import { describe, expect, test } from "bun:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { createDocx, ensureParaIds, FolioDocxReviewer } from "@stll/folio-core/server";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCallUntyped } from "./execute";

const ANCHOR = "Second clause.";

const openReviewer = async () => {
  const model = fromMarkdown(`# Agreement\n\nFirst clause.\n\n${ANCHOR}\n\nTail.`);
  const { docx } = await ensureParaIds(await createDocx(model));
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, { author: "AI" });
  const anchor = reviewer.getContent().find((block) => block.text === ANCHOR)?.id;
  if (anchor === undefined) {
    throw new Error("fixture must expose its anchor block");
  }
  return { reviewer, anchor };
};

const texts = (reviewer: FolioDocxReviewer) => reviewer.getContent().map((block) => block.text);

const suggestInsert = (reviewer: FolioDocxReviewer, anchor: string, styleId: string) =>
  executeFolioToolCallUntyped(
    "suggest_changes",
    { operations: [{ type: "insertAfterBlock", blockId: anchor, text: "Inserted.", styleId }] },
    createReviewerBridge(reviewer, { mode: "tracked-changes" }),
    {},
  );

describe("suggest_changes with a style the document does not define", () => {
  test.each([["NoSuchStyle"], ["TableGrid"]])(
    "is skipped, not applied, and the reviewer saves (%s)",
    async (styleId) => {
      const { reviewer, anchor } = await openReviewer();
      const before = texts(reviewer);
      const out = suggestInsert(reviewer, anchor, styleId);
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
      expect(result.skipped[0]?.reason).toContain("names no paragraph style this document defines");
      expect(result.skipped[0]?.reason).toContain(`"${styleId}"`);
      expect(result.skipped[0]?.reason).toContain("Heading2");
      expect(result.issues).toEqual([
        expect.objectContaining({ code: "missingStyle", recovery: "refreshDocument" }),
      ]);
      expect(texts(reviewer)).toEqual(before);
      await reviewer.toBuffer();
    },
  );

  test("control: a defined paragraph style applies and saves", async () => {
    const { reviewer, anchor } = await openReviewer();
    const out = suggestInsert(reviewer, anchor, "Heading2");
    expect(out.ok).toBe(true);
    expect(out.ok && (out.result as { applied: unknown[] }).applied).toHaveLength(1);
    expect(texts(reviewer)).toContain("Inserted.");
    await reviewer.toBuffer();
  });
});
