import { describe, expect, test } from "bun:test";

import { resolveFloatingTablePageX, resolveFloatingTableX } from "./floatingTablePosition";

const CONTENT = 624;
const TABLE = 333;
const OVERWIDE_TABLE = 700;

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

  test.each([
    ["center", (CONTENT - OVERWIDE_TABLE) / 2],
    ["right", CONTENT - OVERWIDE_TABLE],
    ["outside", CONTENT - OVERWIDE_TABLE],
  ] as const)(
    "preserves the signed offset for an over-wide tblpXSpec=%s table",
    (spec, expected) => {
      expect(resolveFloatingTableX({ tblpXSpec: spec }, undefined, OVERWIDE_TABLE, CONTENT)).toBe(
        expected,
      );
    },
  );

  test("preserves signed over-wide offsets from w:jc fallback", () => {
    expect(resolveFloatingTableX({}, "center", OVERWIDE_TABLE, CONTENT)).toBe(
      (CONTENT - OVERWIDE_TABLE) / 2,
    );
    expect(resolveFloatingTableX({}, "right", OVERWIDE_TABLE, CONTENT)).toBe(
      CONTENT - OVERWIDE_TABLE,
    );
  });
});

describe("resolveFloatingTablePageX", () => {
  test.each([
    ["center", 300],
    ["right", 600],
    ["outside", 600],
  ] as const)("resolves page-anchored tblpXSpec=%s against the page width", (spec, expected) => {
    expect(
      resolveFloatingTablePageX({
        anchor: { horzAnchor: "page", tblpXSpec: spec },
        justification: undefined,
        tableWidth: 200,
        marginWidth: 600,
        pageWidth: 800,
        marginLeft: 40,
        textFrameWidth: 600,
        textFrameLeft: 40,
      }),
    ).toBe(expected);
  });

  test("clamps the final rectangle to the physical page", () => {
    expect(
      resolveFloatingTablePageX({
        anchor: { horzAnchor: "page", tblpX: 760 },
        justification: undefined,
        tableWidth: 200,
        marginWidth: 600,
        pageWidth: 800,
        marginLeft: 40,
        textFrameWidth: 600,
        textFrameLeft: 40,
      }),
    ).toBe(600);
  });

  test("resolves a text-anchored offset from the active column frame", () => {
    expect(
      resolveFloatingTablePageX({
        anchor: { horzAnchor: "text", tblpX: 10 },
        justification: undefined,
        tableWidth: 200,
        marginWidth: 800,
        pageWidth: 1_000,
        marginLeft: 100,
        textFrameWidth: 350,
        textFrameLeft: 550,
      }),
    ).toBe(560);
    expect(
      resolveFloatingTablePageX({
        anchor: { horzAnchor: "text", tblpXSpec: "center" },
        justification: undefined,
        tableWidth: 200,
        marginWidth: 800,
        pageWidth: 1_000,
        marginLeft: 100,
        textFrameWidth: 350,
        textFrameLeft: 550,
      }),
    ).toBe(625);
  });
});
