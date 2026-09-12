import { describe, expect, test } from "bun:test";

import { listRenderingAttrPatch, withDirectListNumbering } from "./listRenderingProjection";

describe("list rendering projection", () => {
  test("uses canonical absence while retaining defined zero values", () => {
    expect(
      listRenderingAttrPatch({
        marker: "",
        level: 0,
        numId: 4,
        isBullet: false,
        startOverride: 0,
      }),
    ).toEqual({ listStartOverride: 0 });
  });

  test("replaces direct numbering without mutating the source formatting", () => {
    const source = { alignment: "both" as const, numPr: { numId: 4, ilvl: 2 } };

    expect(withDirectListNumbering(source, { numId: 4, ilvl: 1 })).toEqual({
      alignment: "both",
      numPr: { numId: 4, ilvl: 1 },
    });
    expect(source.numPr).toEqual({ numId: 4, ilvl: 2 });
  });

  test("clears only direct numbering", () => {
    expect(
      withDirectListNumbering({ alignment: "right", numPr: { numId: 4, ilvl: 2 } }, null),
    ).toEqual({ alignment: "right" });
    expect(withDirectListNumbering({ numPr: { numId: 4, ilvl: 2 } }, null)).toBeUndefined();
  });
});
