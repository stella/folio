import { describe, expect, test } from "bun:test";

import { resolveFloatingTableX } from "./floatingTablePosition";

const CONTENT = 624;
const TABLE = 333;

describe("resolveFloatingTableX", () => {
  test('tblpXSpec supersedes tblpX (§17.4.57: "that value is ignored")', () => {
    expect(
      resolveFloatingTableX({ tblpX: 40, tblpXSpec: "center" }, undefined, TABLE, CONTENT),
    ).toBe((CONTENT - TABLE) / 2);
  });

  test("tblpX applies when no tblpXSpec is authored", () => {
    expect(resolveFloatingTableX({ tblpX: 40 }, undefined, TABLE, CONTENT)).toBe(40);
  });

  test.each([
    ["left", 0],
    ["inside", 0],
    ["center", (CONTENT - TABLE) / 2],
    ["right", CONTENT - TABLE],
    ["outside", CONTENT - TABLE],
  ] as const)("tblpXSpec=%s", (spec, expected) => {
    expect(resolveFloatingTableX({ tblpXSpec: spec }, undefined, TABLE, CONTENT)).toBe(expected);
  });

  test("falls back to w:jc justification, then to the left margin", () => {
    expect(resolveFloatingTableX({}, "center", TABLE, CONTENT)).toBe((CONTENT - TABLE) / 2);
    expect(resolveFloatingTableX({}, "right", TABLE, CONTENT)).toBe(CONTENT - TABLE);
    expect(resolveFloatingTableX({}, undefined, TABLE, CONTENT)).toBe(0);
  });
});
