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
        metadataChanged: 0,
        unchanged: 2,
      },
      stories: [],
      metadataChanges: [],
      privacyReport: { appliedTransforms: [], removedMetadataProperties: [] },
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

  test("renders metadata changes when present", () => {
    const diff: FolioAgentVersionDiff = {
      summaryCounts: {
        added: 0,
        deleted: 0,
        modified: 0,
        formatChanged: 0,
        moved: 0,
        metadataChanged: 2,
        unchanged: 1,
      },
      stories: [],
      changes: [],
      metadataChanges: [
        { property: "title", baseValue: "Initial", revisedValue: "Revised" },
        { property: "creator", baseValue: "Author", revisedValue: null },
      ],
      privacyReport: { appliedTransforms: [], removedMetadataProperties: [] },
    };

    expect(formatVersionDiffForLLM(diff).split("\n")).toEqual([
      "Version diff: 0 added, 0 deleted, 0 modified, 0 format-changed, 0 moved, 2 metadata-changed, 1 unchanged",
      '~ metadata.title: "Initial" -> "Revised"',
      '~ metadata.creator: "Author" -> null',
    ]);
  });

  test("renders applied privacy transforms and their removal report", () => {
    const diff: FolioAgentVersionDiff = {
      summaryCounts: {
        added: 0,
        deleted: 0,
        modified: 0,
        formatChanged: 0,
        moved: 0,
        metadataChanged: 1,
        unchanged: 1,
      },
      stories: [],
      changes: [],
      metadataChanges: [{ property: "revision", baseValue: 1, revisedValue: 2 }],
      privacyReport: {
        appliedTransforms: ["remove-attribution", "remove-timestamps"],
        removedMetadataProperties: ["creator", "modified"],
      },
    };

    expect(formatVersionDiffForLLM(diff).split("\n")).toEqual([
      "Version diff: 0 added, 0 deleted, 0 modified, 0 format-changed, 0 moved, 1 metadata-changed, 1 unchanged",
      "Privacy transforms: remove-attribution, remove-timestamps",
      "Removed metadata fields: creator, modified",
      "~ metadata.revision: 1 -> 2",
    ]);
  });
});
