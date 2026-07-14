import { describe, expect, test } from "bun:test";

import type { FolioAgentVersionDiff } from "./compare";
import { formatVersionDiffForLLM } from "./compare";

const mainHandle = (blockId: string) => ({ story: { type: "main" } as const, blockId });

describe("formatVersionDiffForLLM", () => {
  test("renders summary counts and word-diff markers", () => {
    const diff: FolioAgentVersionDiff = {
      summaryCounts: {
        added: 1,
        deleted: 1,
        modified: 1,
        formatChanged: 1,
        moved: 1,
        unchanged: 2,
      },
      stories: [],
      changes: [
        {
          type: "modified",
          blockId: "00000002",
          kind: "paragraph",
          segments: [
            { type: "equal", text: "Beta " },
            { type: "del", text: "paragraph." },
            { type: "ins", text: "clause." },
          ],
          baseHandle: mainHandle("00000002"),
          revisedHandle: mainHandle("00000002"),
        },
        {
          type: "deleted",
          blockId: "00000003",
          kind: "paragraph",
          text: "Gamma paragraph.",
          baseHandle: mainHandle("00000003"),
        },
        {
          type: "added",
          blockId: "00000006",
          kind: "paragraph",
          text: "Epsilon paragraph.",
          revisedHandle: mainHandle("00000006"),
        },
        {
          type: "formatChanged",
          blockId: "00000007",
          kind: "paragraph",
          text: "Delta paragraph.",
          changedProperties: ["bold", "color"],
          baseHandle: mainHandle("00000007"),
          revisedHandle: mainHandle("00000007"),
        },
        {
          type: "movedFrom",
          blockId: "00000008",
          kind: "paragraph",
          text: "Zeta paragraph moved somewhere else.",
          moveGroupId: 1,
          baseHandle: mainHandle("00000008"),
        },
        {
          type: "movedTo",
          blockId: "00000008",
          kind: "paragraph",
          text: "Zeta paragraph moved somewhere else.",
          moveGroupId: 1,
          revisedHandle: mainHandle("00000008"),
        },
      ],
    };

    expect(formatVersionDiffForLLM(diff).split("\n")).toEqual([
      "Version diff: 1 added, 1 deleted, 1 modified, 1 format-changed, 1 moved, 2 unchanged",
      "~ [00000002] modified: Beta [-paragraph.-]{+clause.+}",
      "- [00000003] deleted: Gamma paragraph.",
      "+ [00000006] added: Epsilon paragraph.",
      "~ [00000007] format changed (bold, color): Delta paragraph.",
      "< [00000008] moved away (move 1): Zeta paragraph moved somewhere else.",
      "> [00000008] moved here (move 1): Zeta paragraph moved somewhere else.",
    ]);
  });
});
