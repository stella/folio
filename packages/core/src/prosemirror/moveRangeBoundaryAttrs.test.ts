import { describe, expect, test } from "bun:test";

import { schema } from "./schema";
import { readMoveRangeBoundaryAttrs } from "./moveRangeBoundaryAttrs";

const read = (marker: unknown) =>
  readMoveRangeBoundaryAttrs(schema.node("moveRangeBoundary", { marker }));

describe("move range boundary attrs", () => {
  test("accepts every serializer-consumed move-start field", () => {
    expect(
      read({
        type: "moveFromRangeStart",
        id: 7,
        name: "move-7",
        author: "Reviewer",
        date: "2026-09-21T00:00:00Z",
        colFirst: 1,
        colLast: 3,
        displacedByCustomXml: "next",
      }),
    ).toEqual({
      ok: true,
      value: {
        type: "moveFromRangeStart",
        id: 7,
        name: "move-7",
        author: "Reviewer",
        date: "2026-09-21T00:00:00Z",
        colFirst: 1,
        colLast: 3,
        displacedByCustomXml: "next",
      },
    });
  });

  test.each([
    [{ type: "moveFromRangeStart", id: 1, name: "n" }, ".author"],
    [{ type: "moveFromRangeStart", id: 1, name: "n", author: "a", date: null }, ".date"],
    [{ type: "moveToRangeStart", id: -1, name: "n", author: "a" }, ".id"],
    [{ type: "moveToRangeStart", id: 1, name: "n", author: "a", colFirst: 1.5 }, ".colFirst"],
    [{ type: "moveToRangeEnd", id: 1, displacedByCustomXml: "sideways" }, ".displacedByCustomXml"],
  ] as const)("rejects malformed marker field %s", (marker, path) => {
    const result = read(marker);

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.issues.some((issue) => issue.path.endsWith(path))).toBe(true);
    }
  });
});
