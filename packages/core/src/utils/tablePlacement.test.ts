import { describe, expect, test } from "bun:test";

import { TABLE_JUSTIFICATION_VALUES } from "../types/documentEnumValues";
import type { TableAlignment } from "../types/formatting";
import { resolveTablePlacementAlignment } from "./tablePlacement";

const CASES = [
  { placement: "left", rightToLeft: false, expected: "left" },
  { placement: "left", rightToLeft: true, expected: "right" },
  { placement: "start", rightToLeft: false, expected: "left" },
  { placement: "start", rightToLeft: true, expected: "right" },
  { placement: "center", rightToLeft: false, expected: "center" },
  { placement: "center", rightToLeft: true, expected: "center" },
  { placement: "right", rightToLeft: false, expected: "right" },
  { placement: "right", rightToLeft: true, expected: "left" },
  { placement: "end", rightToLeft: false, expected: "right" },
  { placement: "end", rightToLeft: true, expected: "left" },
] as const satisfies readonly {
  placement: TableAlignment;
  rightToLeft: boolean;
  expected: "left" | "center" | "right";
}[];

describe("physical table placement", () => {
  test("covers every schema placement", () => {
    expect(new Set(CASES.map(({ placement }) => placement))).toEqual(
      new Set(TABLE_JUSTIFICATION_VALUES),
    );
  });

  test.each(CASES)(
    "resolves $placement to $expected when rightToLeft=$rightToLeft",
    ({ placement, rightToLeft, expected }) => {
      expect(resolveTablePlacementAlignment(placement, rightToLeft)).toBe(expected);
    },
  );
});
