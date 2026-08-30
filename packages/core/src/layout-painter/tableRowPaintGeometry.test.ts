import { describe, expect, test } from "bun:test";

import { ownedRowBottomBorderOffsets } from "./tableRowPaintGeometry";

describe("ownedRowBottomBorderOffsets", () => {
  test("rounds cumulative owned edges from a fractional origin", () => {
    expect(
      ownedRowBottomBorderOffsets({
        origin: 0.25,
        rowHeights: [10.2, 7.7, 5.4],
        snapAfterRow: [true, true],
      }),
    ).toEqual([0.5500000000000007, 0.8500000000000014, 0]);
  });

  test("is stable across fractional origins and three-or-more-row partitions", () => {
    const origins = [0, 0.125, 0.5, 0.875];
    const partitions = [
      [10.2, 7.7, 5.4],
      [1.125, 2.25, 3.375, 4.5],
      [40.01, 0.99, 12.625],
    ];

    for (const origin of origins) {
      for (const rowHeights of partitions) {
        const snapAfterRow = rowHeights.slice(0, -1).map(() => true);
        const once = ownedRowBottomBorderOffsets({ origin, rowHeights, snapAfterRow });
        const twice = ownedRowBottomBorderOffsets({ origin, rowHeights, snapAfterRow });

        expect(twice).toEqual(once);
        expect(once.at(-1)).toBe(0);
        expect(once.every((offset) => offset >= 0 && offset < 1)).toBe(true);
      }
    }
  });

  test("leaves unowned and final edges unchanged", () => {
    expect(
      ownedRowBottomBorderOffsets({
        origin: 0.25,
        rowHeights: [10.2, 7.7, 5.4],
        snapAfterRow: [false, true],
      }),
    ).toEqual([0, 0.8500000000000014, 0]);
  });
});
