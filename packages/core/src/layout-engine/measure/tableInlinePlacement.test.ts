import { describe, expect, test } from "bun:test";

import type { TableBlock } from "../types";
import { resolveTableInlinePlacement } from "./tableInlinePlacement";

const tableWithLeadingPadding = (bidi: boolean, padding: number): TableBlock => ({
  kind: "table",
  id: "table",
  bidi,
  indent: 12,
  rows: [
    {
      id: "row",
      cells: [
        {
          id: "cell",
          blocks: [],
          padding: { top: 0, right: padding, bottom: 0, left: padding },
        },
      ],
    },
  ],
});

describe("resolveTableInlinePlacement", () => {
  test.each([
    { bidi: false, alignment: "left" as const },
    { bidi: true, alignment: "right" as const },
  ])("anchors a $alignment table border independently of cell padding", ({ bidi, alignment }) => {
    for (const padding of [0, 7, 12, 96]) {
      expect(resolveTableInlinePlacement(tableWithLeadingPadding(bidi, padding))).toEqual({
        alignment,
        offset: 12,
      });
    }
  });

  test.each([
    { bidi: false, justification: "left" as const, alignment: "left" as const, offset: 12 },
    { bidi: false, justification: "right" as const, alignment: "right" as const, offset: 0 },
    { bidi: true, justification: "left" as const, alignment: "right" as const, offset: 12 },
    { bidi: true, justification: "right" as const, alignment: "left" as const, offset: 0 },
  ])(
    "maps $justification to physical $alignment when bidi=$bidi",
    ({ bidi, justification, alignment, offset }) => {
      const table = tableWithLeadingPadding(bidi, 7);
      table.justification = justification;

      expect(resolveTableInlinePlacement(table)).toEqual({ alignment, offset });
    },
  );

  test("lets a row justification override the table before applying the indent", () => {
    const table = tableWithLeadingPadding(false, 7);
    table.justification = "left";

    expect(resolveTableInlinePlacement(table, "right")).toEqual({
      alignment: "right",
      offset: 0,
    });
  });
});

describe("resolveTableInlinePlacement with a text-edge indent", () => {
  const legacyTable = (bidi: boolean, padding: number): TableBlock => {
    const table = tableWithLeadingPadding(bidi, padding);
    table.indentCompatibility = { type: "legacy" };
    return table;
  };

  test.each([
    { bidi: false, alignment: "left" as const },
    { bidi: true, alignment: "right" as const },
  ])("pulls the $alignment border back by the leading cell margin", ({ bidi, alignment }) => {
    for (const padding of [0, 7, 12, 96]) {
      expect(resolveTableInlinePlacement(legacyTable(bidi, padding))).toEqual({
        alignment,
        offset: 12 - padding,
      });
    }
  });

  test("moves an unindented table border out by the default leading cell margin", () => {
    const table = legacyTable(false, 7);
    delete table.indent;

    expect(resolveTableInlinePlacement(table)).toEqual({ alignment: "left", offset: -7 });
  });

  test("leaves the trailing edge of a right-justified table alone", () => {
    const table = legacyTable(false, 7);
    table.justification = "right";

    expect(resolveTableInlinePlacement(table)).toEqual({ alignment: "right", offset: 0 });
  });

  test("leaves a centered table alone", () => {
    const table = legacyTable(false, 7);
    table.justification = "center";

    expect(resolveTableInlinePlacement(table)).toEqual({ alignment: "center" });
  });

  test("falls back to the default cell margin when the table has no rows", () => {
    expect(
      resolveTableInlinePlacement({
        kind: "table",
        id: "empty",
        rows: [],
        indent: 12,
        indentCompatibility: { type: "legacy" },
      }),
    ).toEqual({ alignment: "left", offset: 5 });
  });
});
