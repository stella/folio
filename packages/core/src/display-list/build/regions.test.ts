/**
 * The tree has to index the paint list exactly.
 *
 * A region that claimed primitives it does not own, or that left a gap in the
 * middle of a line, would make a backend paint the page in an order the
 * producer did not choose, and would put a click in the wrong place. Both
 * properties are cheap to state and neither is obvious from reading the walk.
 */

import { describe, expect, test } from "bun:test";

import { createPageComposer } from "./regions";
import { walkRegions } from "../primitives";
import { BLACK } from "../primitives";
import type { DisplayHitRegion, DisplayPrimitive } from "../types";

const rect = (xPx: number): DisplayRectLike => ({ xPx, yPx: 0, widthPx: 10, heightPx: 10 });

type DisplayRectLike = { xPx: number; yPx: number; widthPx: number; heightPx: number };

const primitive = (xPx: number): DisplayPrimitive => ({
  kind: "rect",
  rect: rect(xPx),
  fill: BLACK,
});

/** Every primitive index a region tree claims, in the order a walk yields it. */
const walked = (
  primitives: readonly DisplayPrimitive[],
  regions: readonly DisplayHitRegion[],
): { order: number[]; entered: string[] } => {
  const order: number[] = [];
  const entered: string[] = [];
  walkRegions(primitives, regions, {
    onPrimitive: (_, index) => order.push(index),
    enter: (region) => entered.push(`+${region.kind}`),
    exit: (region) => entered.push(`-${region.kind}`),
  });
  return { order, entered };
};

describe("a page's regions index its primitives", () => {
  test("a region owns exactly what was painted inside it", () => {
    const composer = createPageComposer();
    composer.push([primitive(0)]);
    composer.region({ kind: "paragraph", rect: rect(1) }, () => {
      composer.push([primitive(1)]);
      composer.region({ kind: "line", rect: rect(2) }, () => {
        composer.push([primitive(2), primitive(3)]);
      });
      composer.push([primitive(4)]);
    });
    composer.push([primitive(5)]);

    const regions = composer.regions();
    expect(regions).toHaveLength(1);
    expect(regions.at(0)).toMatchObject({ kind: "paragraph", from: 1, to: 5 });
    expect(regions.at(0)?.children.at(0)).toMatchObject({ kind: "line", from: 2, to: 4 });
  });

  test("a walk yields every primitive once, in paint order", () => {
    const composer = createPageComposer();
    composer.push([primitive(0)]);
    composer.region({ kind: "paragraph", rect: rect(1) }, () => {
      composer.region({ kind: "line", rect: rect(1) }, () => {
        composer.push([primitive(1)]);
      });
      composer.region({ kind: "line", rect: rect(2) }, () => {
        composer.push([primitive(2)]);
      });
    });
    composer.push([primitive(3)]);

    const { order, entered } = walked(composer.primitives(), composer.regions());
    expect(order).toEqual([0, 1, 2, 3]);
    expect(entered).toEqual(["+paragraph", "+line", "-line", "+line", "-line", "-paragraph"]);
  });

  test("a region that painted nothing still bounds a click", () => {
    const composer = createPageComposer();
    composer.region({ kind: "emptyRun", rect: rect(0) }, () => {
      // A paragraph with no text paints nothing and still takes a caret.
    });

    expect(composer.regions().at(0)).toMatchObject({ kind: "emptyRun", from: 0, to: 0 });
    expect(walked(composer.primitives(), composer.regions()).order).toEqual([]);
  });

  test("a child's range lies inside its parent's, and siblings do not overlap", () => {
    const composer = createPageComposer();
    composer.region({ kind: "table", rect: rect(0) }, () => {
      for (const row of [0, 1]) {
        composer.region({ kind: "tableRow", rect: rect(row) }, () => {
          composer.region({ kind: "tableCell", rect: rect(row) }, () => {
            composer.push([primitive(row)]);
          });
        });
      }
    });

    const check = (region: DisplayHitRegion): void => {
      let previousEnd = region.from;
      for (const child of region.children) {
        expect(child.from).toBeGreaterThanOrEqual(previousEnd);
        expect(child.to).toBeLessThanOrEqual(region.to);
        previousEnd = child.to;
        check(child);
      }
    };
    const table = composer.regions().at(0);
    expect(table).toBeDefined();
    if (table) {
      check(table);
    }
  });
});
