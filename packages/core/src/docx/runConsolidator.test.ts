import { describe, expect, test } from "bun:test";
import type { Run } from "../types/document";
import { consolidateRuns } from "./runConsolidator";

const changedRun = (text: string): Run => ({
  type: "run",
  formatting: { fontSizeCs: 22 },
  propertyChanges: [
    {
      type: "runPropertyChange",
      info: {
        id: 1,
        author: "Reviewer",
        date: "2026-09-08T08:00:00Z",
      },
      previousFormatting: { bold: true, fontSizeCs: 22 },
      currentFormatting: { fontSizeCs: 22 },
    },
  ],
  content: [{ type: "text", text }],
});

describe("run consolidation boundaries", () => {
  test("preserves both boundaries of a run-property revision", () => {
    const unchanged = (text: string): Run => ({
      type: "run",
      formatting: { fontSizeCs: 22 },
      content: [{ type: "text", text }],
    });

    const result = consolidateRuns([
      unchanged("before"),
      changedRun("changed"),
      unchanged("after"),
    ]);

    expect(result).toHaveLength(3);
    expect(result.at(1)).toEqual(changedRun("changed"));
    expect(result.map((run) => run.content.at(0))).toEqual([
      { type: "text", text: "before" },
      { type: "text", text: "changed" },
      { type: "text", text: "after" },
    ]);
  });
});
