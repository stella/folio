import { describe, expect, test } from "bun:test";

import { parseShapeGeometryAdjustments } from "./shapeGeometryAdjustments";

describe("parseShapeGeometryAdjustments", () => {
  test("accepts bounded named formulas", () => {
    expect(
      parseShapeGeometryAdjustments(JSON.stringify([{ name: "adj", formula: "val 25000" }])),
    ).toEqual([{ name: "adj", formula: "val 25000" }]);
  });

  test.each([
    "not json",
    "{}",
    JSON.stringify([{ name: "adj" }]),
    JSON.stringify([{ name: "", formula: "val 1" }]),
    JSON.stringify(Array.from({ length: 33 }, () => ({ name: "adj", formula: "val 1" }))),
  ])("rejects malformed or unbounded input", (raw) => {
    expect(parseShapeGeometryAdjustments(raw)).toBeUndefined();
  });
});
