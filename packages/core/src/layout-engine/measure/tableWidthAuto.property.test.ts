/**
 * `w:tblW`/`w:tcW` pair a number with a `ST_TblWidth` type, and only `dxa` and
 * `pct` make the number a measurement. Under `auto` (§17.18.87) the consumer
 * sizes the table to its content and under `nil` there is no width, so `w:w` is
 * whatever the producer happened to leave there. Two of the three resolvers
 * lumped `auto` in with `dxa` and measured that leftover, pinning a table that
 * should autofit; a third, correct one had no callers at all.
 *
 * `resolveTableWidthPx` is now the only resolver. The property holds over the
 * whole `ST_TblWidth` enum and over both measurement paths.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import type { TableBlock, TableWidthType } from "../types";
import { resolveTableWidthPx } from "../types";
import { measureTableBlock } from "./measureBlocks";

const TABLE_WIDTH_TYPES = [
  "auto",
  "dxa",
  "nil",
  "pct",
] as const satisfies readonly TableWidthType[];

const CONTENT_WIDTH = 600;
/** A `w:w` big enough that measuring it would visibly overrun the content box. */
const LEFTOVER_TWIPS = 9360;

const tableBlock = (width: number, widthType: TableWidthType): TableBlock => ({
  kind: "table",
  id: 1 as TableBlock["id"],
  width,
  widthType,
  columnWidths: [200, 200],
  rows: [
    {
      cells: [
        { blocks: [], colSpan: 1, rowSpan: 1 },
        { blocks: [], colSpan: 1, rowSpan: 1 },
      ],
    } as unknown as TableBlock["rows"][number],
  ],
});

describe("ST_TblWidth auto", () => {
  test('w:type="auto" leaves the table at its content width instead of pinning it', () => {
    const autoFit = measureTableBlock(tableBlock(LEFTOVER_TWIPS, "auto"), CONTENT_WIDTH);
    const pinned = measureTableBlock(tableBlock(LEFTOVER_TWIPS, "dxa"), CONTENT_WIDTH);

    expect(autoFit.totalWidth).toBeLessThanOrEqual(CONTENT_WIDTH);
    expect(pinned.totalWidth).toBeGreaterThan(CONTENT_WIDTH);
  });

  test(
    "only dxa and pct resolve to a length",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom(...TABLE_WIDTH_TYPES),
          fc.integer({ min: 1, max: 20_000 }),
          (widthType, width) => {
            const resolved = resolveTableWidthPx(width, widthType, CONTENT_WIDTH);

            if (widthType === "auto" || widthType === "nil") {
              expect(resolved).toBeUndefined();
              return;
            }
            expect(resolved).toBeGreaterThan(0);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "measurement agrees with the resolver across the enum",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...TABLE_WIDTH_TYPES), (widthType) => {
          const block = tableBlock(LEFTOVER_TWIPS, widthType);
          const resolved = resolveTableWidthPx(block.width, widthType, CONTENT_WIDTH);
          const measured = measureTableBlock(block, CONTENT_WIDTH);

          if (resolved === undefined) {
            expect(measured.totalWidth).toBeLessThanOrEqual(CONTENT_WIDTH);
            return;
          }
          expect(measured.totalWidth).toBeCloseTo(resolved, 5);
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );
});
