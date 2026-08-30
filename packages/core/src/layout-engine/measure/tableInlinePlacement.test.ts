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
});
