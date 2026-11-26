import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import type { StepMap } from "prosemirror-transform";

import { markupViewPositionMap } from "./markupViewPositionMap";
import { schema } from "./schema";

const paragraph = (content: PMNode[], paraId?: string) =>
  schema.node("paragraph", paraId ? { paraId } : null, content);
const doc = (content: PMNode[]) => schema.node("doc", null, content);

const rangesOf = (map: StepMap): number[][] => {
  const ranges: number[][] = [];
  map.forEach((from, to, mappedFrom, mappedTo) =>
    ranges.push([from, to, mappedFrom, mappedTo]),
  );
  return ranges;
};

const revision = (kind: "deletion" | "insertion", id: number) =>
  schema.mark(kind, { revisionId: id, author: "Reviewer" });

describe("authored positions of resolved markup views", () => {
  test.each([
    { decision: "accept", hidden: "deletion" },
    { decision: "reject", hidden: "insertion" },
  ] as const)(
    "keeps the visible occurrence of equal text",
    ({ decision, hidden }) => {
      const before = doc([
        paragraph([
          schema.text("echo", [revision(hidden, 1)]),
          schema.text("echo"),
        ]),
      ]);
      const after = doc([paragraph([schema.text("echo")])]);
      const map = markupViewPositionMap({ before, after, decision });
      expect(rangesOf(map)).toEqual([[1, 5, 1, 1]]);
      expect(map.mapResult(3).deletedAcross).toBe(true);
      expect(map.invert().map(1, 1)).toBe(5);
      expect(map.invert().map(4, 1)).toBe(8);
    },
  );

  test("retains pending text still present in a partial model result", () => {
    const retained = schema.text("echo", [revision("deletion", 2)]);
    const before = doc([
      paragraph([
        schema.text("echo", [revision("deletion", 1)]),
        schema.text("echo"),
        retained,
      ]),
    ]);
    const after = doc([paragraph([schema.text("echo"), retained])]);
    const map = markupViewPositionMap({ before, after, decision: "accept" });
    expect(rangesOf(map)).toEqual([[1, 5, 1, 1]]);
    expect(map.invert().map(8, 1)).toBe(12);
  });

  test("anchors a paragraph join at the surviving paragraph's end", () => {
    const before = doc([
      paragraph([schema.text("one")], "00000001"),
      paragraph([schema.text("two")], "00000002"),
      paragraph([schema.text("standing")], "00000003"),
    ]);
    const after = doc([
      paragraph([schema.text("onetwo")], "00000002"),
      paragraph([schema.text("standing")], "00000003"),
    ]);
    const map = markupViewPositionMap({ before, after, decision: "accept" });
    expect(rangesOf(map)).toEqual([[4, 6, 4, 4]]);
    expect(map.invert().map(4, -1)).toBe(4);
    expect(map.invert().map(4, 1)).toBe(6);
    expect(map.map(13)).toBe(11);
  });

  test("distinguishes identical paragraphs by the surviving paragraph ID", () => {
    const gone = paragraph([schema.text("echo")], "00000001");
    const kept = paragraph([schema.text("echo")], "00000002");
    const map = markupViewPositionMap({
      before: doc([gone, kept]),
      after: doc([kept]),
      decision: "accept",
    });
    expect(rangesOf(map)).toEqual([[0, 6, 0, 0]]);
    expect(map.invert().map(1, 1)).toBe(7);
  });

  test("preserves both sides of a join beyond the alignment budget", () => {
    const left = "left ".repeat(1_024);
    const right = "right ".repeat(1_024);
    const before = doc([
      paragraph([schema.text(left)], "00000001"),
      paragraph([schema.text(right)], "00000002"),
    ]);
    const after = doc([paragraph([schema.text(left + right)], "00000002")]);
    const map = markupViewPositionMap({ before, after, decision: "reject" });
    expect(rangesOf(map)).toEqual([
      [left.length + 1, left.length + 3, left.length + 1, left.length + 1],
    ]);
    expect(map.invert().map(left.length + 1, 1)).toBe(left.length + 3);
  });

  test("preserves long standing text when a join retires an exhausted comment", () => {
    const standing = "standing text ".repeat(32);
    const keptReference = schema.node("commentReference", { commentId: 0 });
    const before = doc([
      paragraph([schema.text(standing), keptReference], "00000001"),
      paragraph(
        [
          schema.text("removed text ".repeat(32), [revision("insertion", 1)]),
          schema.node("commentReference", { commentId: 1 }),
        ],
        "00000002",
      ),
    ]);
    const after = doc([
      paragraph([schema.text(standing), keptReference], "00000002"),
    ]);
    const map = markupViewPositionMap({ before, after, decision: "reject" });
    const afterText = standing.length + 1;
    expect(map.invert().map(1, 1)).toBe(1);
    expect(map.invert().map(afterText, -1)).toBe(afterText);
    expect(map.invert().map(afterText + 1, -1)).toBe(afterText + 1);
  });

  test("maps each invisible gap in a long paragraph", () => {
    const standing = "x".repeat(40_000);
    const before = doc([
      paragraph([
        schema.text(standing),
        schema.text("first hidden", [revision("deletion", 1)]),
        schema.text(standing),
        schema.text("next hidden", [revision("deletion", 2)]),
        schema.text(standing),
      ]),
    ]);
    const after = doc([paragraph([schema.text(standing.repeat(3))])]);
    const map = markupViewPositionMap({ before, after, decision: "accept" });
    expect(rangesOf(map)).toEqual([
      [40_001, 40_013, 40_001, 40_001],
      [80_013, 80_024, 80_001, 80_001],
    ]);
  });

  test("maps standing text beside a resolved move with four range boundaries", () => {
    const standing = "x".repeat(1_024);
    const boundary = (type: string, id: number) =>
      schema.node("moveRangeBoundary", {
        marker: {
          type,
          id,
          name: "move1",
          author: "Reviewer",
          date: "2026-01-01T00:00:00Z",
        },
      });
    const before = doc([
      paragraph([
        schema.text(standing),
        boundary("moveFromRangeStart", 1),
        schema.text("move", [
          schema.mark("deletion", {
            revisionId: 1,
            moveKind: "moveFrom",
            author: "Reviewer",
          }),
        ]),
        boundary("moveFromRangeEnd", 1),
        boundary("moveToRangeStart", 2),
        schema.text("move", [
          schema.mark("insertion", {
            revisionId: 2,
            moveKind: "moveTo",
            author: "Reviewer",
          }),
        ]),
        boundary("moveToRangeEnd", 2),
      ]),
    ]);
    const after = doc([paragraph([schema.text(`${standing}move`)])]);
    const map = markupViewPositionMap({ before, after, decision: "accept" });
    expect(rangesOf(map)).toEqual([
      [1025, 1032, 1025, 1025],
      [1036, 1037, 1029, 1029],
    ]);
    expect(map.invert().map(1025, 1)).toBe(1032);
  });

  test("keeps text addresses when resolution recombines move boundaries", () => {
    const standing = "x".repeat(1_024);
    const from = schema.node("moveRangeBoundary", {
      marker: { type: "moveFromRangeEnd", id: 1 },
    });
    const to = schema.node("moveRangeBoundary", {
      marker: {
        type: "moveToRangeStart",
        id: 2,
        name: "move",
        author: "Reviewer",
      },
    });
    const before = doc([
      paragraph([
        schema.text(standing),
        schema.text("move", [revision("deletion", 1)]),
        from,
        to,
        schema.text("move", [revision("insertion", 2)]),
      ]),
    ]);
    const after = doc([
      paragraph([
        schema.text(standing),
        schema.node("rangeAnchor", {
          markers: [{ type: "moveFromRangeEnd", id: 1 }],
        }),
        to,
        schema.text("move"),
      ]),
    ]);
    const map = markupViewPositionMap({ before, after, decision: "accept" });
    expect(rangesOf(map)).toEqual([[1025, 1030, 1025, 1026]]);
    expect(map.invert().map(1027, 1)).toBe(1031);
  });

  test("keeps the opening of the first standing paragraph between two cuts", () => {
    const gone = schema.node("table", null, [
      schema.node(
        "tableRow",
        {
          trDel: { revisionId: 1, author: "Reviewer" },
        },
        [
          schema.node("tableCell", null, [
            paragraph([schema.text("discarded")]),
          ]),
        ],
      ),
    ]);
    const to = schema.node("moveRangeBoundary", {
      marker: {
        type: "moveToRangeStart",
        id: 2,
        name: "move",
        author: "Reviewer",
      },
    });
    const after = doc([
      paragraph(
        [
          schema.node("rangeAnchor", {
            markers: [{ type: "moveFromRangeEnd", id: 1 }],
          }),
          to,
          schema.text("move"),
        ],
        "00000001",
      ),
    ]);
    const deleted = "delete ".repeat(4);
    const before = doc([
      gone,
      paragraph(
        [
          schema.text(deleted, [revision("deletion", 3)]),
          to,
          schema.text("move", [revision("insertion", 2)]),
        ],
        "00000001",
      ),
    ]);
    const map = markupViewPositionMap({ before, after, decision: "accept" });
    expect(rangesOf(map)).toEqual([
      [0, gone.nodeSize, 0, 0],
      [gone.nodeSize + 1, gone.nodeSize + 1 + deleted.length, 1, 2],
    ]);
    expect(map.invert().map(0, 1)).toBe(gone.nodeSize);
    expect(map.invert().map(1, -1)).toBe(gone.nodeSize + 1);
  });
});
