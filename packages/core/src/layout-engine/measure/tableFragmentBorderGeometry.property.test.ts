import { expect, test } from "bun:test";
import fc from "fast-check";

import { tableFragmentBottomBorders } from "./tableFragmentBorderGeometry";

test("a split row projects every visible bottom border onto its fragment edge", () => {
  fc.assert(
    fc.property(
      fc.array(
        fc.record({
          width: fc.integer({ min: 8, max: 160 }),
          borderWidth: fc.integer({ min: 0, max: 8 }),
          style: fc.constantFrom("solid", "dashed", "dotted", "double", "none", "nil"),
        }),
        { minLength: 1, maxLength: 8 },
      ),
      (columns) => {
        const cells = columns.map(({ borderWidth, style }, index) => ({
          id: `cell-${String(index)}`,
          blocks: [],
          borders: { bottom: { width: borderWidth, style, color: "#123456" } },
        }));
        const columnWidths = columns.map(({ width }) => width);
        const block = {
          kind: "table",
          id: "table",
          rows: [{ id: "row", cells }],
          columnWidths,
        } as const;
        const measure = {
          kind: "table",
          rows: [
            {
              cells: columnWidths.map((width) => ({ blocks: [], width, height: 100 })),
              height: 100,
            },
          ],
          columnWidths,
          totalWidth: columnWidths.reduce((total, width) => total + width, 0),
          totalHeight: 100,
        } as const;
        const fragment = {
          kind: "table",
          blockId: "table",
          x: 0,
          y: 0,
          width: measure.totalWidth,
          height: 40,
          fromRow: 0,
          toRow: 1,
          bottomClip: 40,
          continuesOnNext: true,
        } as const;

        const borders = tableFragmentBottomBorders({ fragment, block, measure });
        let left = 0;
        const expected = columns.flatMap(({ width, style }) => {
          const segment = style === "none" || style === "nil" ? [] : [{ left, width, style }];
          left += width;
          return segment;
        });

        expect(
          borders.map(({ left: segmentLeft, width, border }) => ({
            left: segmentLeft,
            width,
            style: border.style,
          })),
        ).toEqual(expected);
      },
    ),
  );
});
