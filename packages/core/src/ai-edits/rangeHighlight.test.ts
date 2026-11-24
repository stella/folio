import { expect, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { TextFormatting } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";

const blockId = "10000000";
const text = "Before target after";

type HighlightOptions = {
  highlight: NonNullable<TextFormatting["highlight"]> | null;
  mode: "direct" | "tracked-changes" | "suggested";
};

const setHighlight = (
  reviewer: FolioDocxReviewer,
  { highlight, mode }: HighlightOptions,
) => {
  const range = createFolioAITextRangeHandle({
    blockId,
    text,
    startOffset: 7,
    endOffset: 13,
  });
  if (!range) throw new Error("Expected a text range");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [
        { id: "format", type: "formatRange", range, formatting: { highlight } },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(1);
  return result;
};

const runHighlights = (reviewer: FolioDocxReviewer) => {
  const paragraph = reviewer.toDocument().package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  const spans: { text: string; highlight: TextFormatting["highlight"] }[] = [];
  for (const item of paragraph.content) {
    if (item.type !== "run") continue;
    const runText = item.content
      .flatMap((child) => (child.type === "text" ? [child.text] : []))
      .join("");
    if (runText.length === 0) continue;
    const highlight = item.formatting?.highlight;
    const last = spans.at(-1);
    if (last && last.highlight === highlight) last.text += runText;
    else spans.push({ text: runText, highlight });
  }
  return spans;
};

for (const session of ["canonical"] as const) {
  const reopen = async (reviewer: FolioDocxReviewer) =>
    FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { session });

  test(`${session} sets and clears a direct highlight only inside the range`, async () => {
    let reviewer = await FolioDocxReviewer.fromBuffer(
      await paragraphsDocx([[textRun(text, { highlight: "yellow" })]]),
      { session },
    );
    for (const value of ["darkGreen", "none", null] as const) {
      setHighlight(reviewer, { highlight: value, mode: "direct" });
      reviewer = await reopen(reviewer);
      expect(runHighlights(reviewer)).toEqual([
        { text: "Before ", highlight: "yellow" },
        { text: "target", highlight: value ?? undefined },
        { text: " after", highlight: "yellow" },
      ]);
    }
  });

  for (const mode of ["tracked-changes", "suggested"] as const)
    test(`${session} preserves the original highlight when rejecting ${mode}`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await paragraphsDocx([[textRun(text, { highlight: "yellow" })]]),
        { session },
      );
      const result = setHighlight(reviewer, { highlight: "cyan", mode });
      if (mode === "suggested") {
        const suggestionId = result.applied.at(0)?.suggestionId;
        if (!suggestionId) throw new Error("Expected a suggestion");
        expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
      }
      const accepted = await reopen(reviewer);
      expect(accepted.acceptAll().resolved).toBeGreaterThan(0);
      expect(runHighlights(await reopen(accepted))).toEqual([
        { text: "Before ", highlight: "yellow" },
        { text: "target", highlight: "cyan" },
        { text: " after", highlight: "yellow" },
      ]);

      const rejected = await reopen(reviewer);
      expect(rejected.rejectAll().resolved).toBeGreaterThan(0);
      expect(runHighlights(await reopen(rejected))).toEqual([
        { text, highlight: "yellow" },
      ]);
    });
}
