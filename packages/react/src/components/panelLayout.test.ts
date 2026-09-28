import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  COMMENTS_TRACK_WIDTH,
  PANEL_METRICS,
  computePanelLayout,
  type CommentsPresentation,
  type OutlinePresentation,
  type PanelLayoutInput,
} from "./panelLayout";

const input = fc.record({
  availableWidth: fc.integer({ min: 0, max: 4000 }),
  pageWidth: fc.integer({ min: 100, max: 2400 }),
  outline: fc.constantFrom("absent", "available"),
  comments: fc.constantFrom("closed", "open"),
}) satisfies fc.Arbitrary<PanelLayoutInput>;

/** Presentations from least to most room taken beside the page. */
const OUTLINE_RANK = { none: 0, drawer: 1, rail: 2, column: 3 } as const satisfies Record<
  OutlinePresentation,
  number
>;
const COMMENTS_RANK = { hidden: 0, drawer: 1, column: 2 } as const satisfies Record<
  CommentsPresentation,
  number
>;

describe("computePanelLayout", () => {
  test("the tracks it hands out always fit beside the page", () => {
    fc.assert(
      fc.property(input, (layoutInput) => {
        const layout = computePanelLayout(layoutInput);
        const taken =
          layout.outlineTrackWidth +
          layoutInput.pageWidth +
          2 * PANEL_METRICS.pageMargin +
          layout.commentsGutter;
        const anyTrack = layout.outlineTrackWidth > 0 || layout.commentsGutter > 0;
        expect(!anyTrack || taken <= layoutInput.availableWidth).toBe(true);
        expect(layout.commentsGutter).toBe(layout.comments === "column" ? COMMENTS_TRACK_WIDTH : 0);
      }),
    );
  });

  test("a panel the user did not ask for takes no room", () => {
    fc.assert(
      fc.property(input, (layoutInput) => {
        const layout = computePanelLayout(layoutInput);
        expect(layout.outline === "none").toBe(layoutInput.outline === "absent");
        expect(layout.comments === "hidden").toBe(layoutInput.comments === "closed");
      }),
    );
  });

  test("more width never demotes a panel", () => {
    const monotonicLayout = fc.property(
      input,
      fc.integer({ min: 0, max: 2000 }),
      (layoutInput, extra) => {
        const narrower = computePanelLayout(layoutInput);
        const wider = computePanelLayout({
          ...layoutInput,
          availableWidth: layoutInput.availableWidth + extra,
        });
        expect(OUTLINE_RANK[wider.outline]).toBeGreaterThanOrEqual(OUTLINE_RANK[narrower.outline]);
        expect(COMMENTS_RANK[wider.comments]).toBeGreaterThanOrEqual(
          COMMENTS_RANK[narrower.comments],
        );
      },
    );
    fc.assert(monotonicLayout, {
      seed: -448549325,
      path: "66:5:0:0:0:0:1:0:6:1:3:4:4:4:4:3",
    });
    fc.assert(monotonicLayout);
  });

  test("the tier names the presentation", () => {
    fc.assert(
      fc.property(input, (layoutInput) => {
        const layout = computePanelLayout(layoutInput);
        const hasDrawer = layout.outline === "drawer" || layout.comments === "drawer";
        if (hasDrawer) {
          expect(layout.tier).toBe("narrow");
        } else if (layout.outline === "rail") {
          expect(layout.tier).toBe("medium");
        } else {
          expect(layout.tier).toBe("wide");
        }
        // At a tier's threshold, the tier fits.
        const atWide = computePanelLayout({
          ...layoutInput,
          availableWidth: layout.thresholds.wide,
        });
        expect(atWide.tier).toBe("wide");
      }),
    );
  });
});
