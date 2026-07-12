import { describe, expect, test } from "bun:test";

import type { FolioAgentVersionDiff } from "./compare";
import { formatVersionDiffForLLM } from "./compare";

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
        },
        {
          type: "deleted",
          blockId: "00000003",
          kind: "paragraph",
          text: "Gamma paragraph.",
        },
        {
          type: "added",
          blockId: "00000006",
          kind: "paragraph",
          text: "Epsilon paragraph.",
        },
        {
          type: "formatChanged",
          blockId: "00000007",
          kind: "paragraph",
          text: "Delta paragraph.",
          changedProperties: ["bold", "color"],
        },
        {
          type: "movedFrom",
          blockId: "00000008",
          kind: "paragraph",
          text: "Zeta paragraph moved somewhere else.",
          moveGroupId: 1,
        },
        {
          type: "movedTo",
          blockId: "00000008",
          kind: "paragraph",
          text: "Zeta paragraph moved somewhere else.",
          moveGroupId: 1,
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
