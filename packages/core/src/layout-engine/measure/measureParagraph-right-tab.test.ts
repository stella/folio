import { describe, expect, test } from "bun:test";

import { calculateTabWidth } from "../../prosemirror/utils/tabCalculator";
import type { TabContext } from "../../prosemirror/utils/tabCalculator";
import { withFakeTextMeasure } from "./__tests__/fakeTextMeasure";
import { measureParagraph } from "./measureParagraph";

// Regression: the legacy `computeTabWidth` in measureParagraph ignored the
// tab stop's `val` field, so a right (`end`) or center stop was measured as
// `stopPx + followingTextWidth`. The painter already right-/center-anchors
// such content via `calculateTabWidth`, so the wrap fired in the measurer
// but not in the painter. See eigenpal #576.
describe("measureParagraph — right/center tab stops (eigenpal #576)", () => {
  test("leading left tab does not shrink to make following text fit the line", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "leading-left-tab-wrap",
          runs: [{ kind: "tab" }, { kind: "text", text: "abcdefghijklmno", fontSize: 11 }],
          attrs: {
            tabs: [{ val: "start", pos: 1500 }],
          },
        },
        150,
      );

      expect(measure.lines).toHaveLength(2);
      expect(measure.lines[0]?.toRun).toBe(0);
      expect(measure.lines[1]?.fromRun).toBe(1);
    });
  });

  test("non-leading left tab may shrink to keep tabbed label text on the line", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "label-left-tab",
          runs: [
            { kind: "text", text: "(a)", fontSize: 11 },
            { kind: "tab" },
            { kind: "text", text: "abcdefghijklmno", fontSize: 11 },
          ],
          attrs: {
            tabs: [{ val: "start", pos: 1500 }],
          },
        },
        150,
      );

      expect(measure.lines).toHaveLength(1);
    });
  });

  test("non-leading left tab keeps its stop before multi-line prose", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "multi-line-label-left-tab",
          runs: [
            { kind: "text", text: "(i)", fontSize: 11 },
            { kind: "tab" },
            { kind: "text", text: "word ".repeat(12), fontSize: 11 },
          ],
          attrs: {
            tabs: [{ val: "start", pos: 1500 }],
          },
        },
        150,
      );

      expect(measure.lines.length).toBeGreaterThan(1);
      expect(measure.lines[0]?.toRun).toBe(2);
      expect(measure.lines[0]?.width).toBeGreaterThan(100);
    });
  });

  test("right tab followed by short text does not wrap", () => {
    withFakeTextMeasure(() => {
      // Right tab stop at 5000 twips ≈ 333.33px. With the bug, the measurer
      // advances to the stop unconditionally (tab ≈ 308.33px on top of the
      // 25px title), then the trailing "page" run (20px) overflows the
      // 340px line and starts a new one. After the fix, end-alignment
      // subtracts the trailing width: the tab pulls "page" right of the
      // stop and the whole line fits.
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "right-tab",
          runs: [
            { kind: "text", text: "Title", fontSize: 11 },
            { kind: "tab" },
            { kind: "text", text: "page", fontSize: 11 },
          ],
          attrs: {
            tabs: [{ val: "end", pos: 5000 }],
          },
        },
        340,
      );

      expect(measure.lines).toHaveLength(1);
    });
  });

  test("center tab followed by text does not wrap", () => {
    withFakeTextMeasure(() => {
      // Center stop at 5000 twips ≈ 333.33px. "center" is 30px, so a
      // center-aligned anchor needs only `tabWidth = (333.33 - 20) - 15`,
      // putting the line at 333.33 + 15 = 348.33px — fits in 350. The
      // legacy measurer would report 20 + 313.33 + 30 = 363.33px and wrap.
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "center-tab",
          runs: [
            { kind: "text", text: "left", fontSize: 11 },
            { kind: "tab" },
            { kind: "text", text: "center", fontSize: 11 },
          ],
          attrs: {
            tabs: [{ val: "center", pos: 5000 }],
          },
        },
        350,
      );

      expect(measure.lines).toHaveLength(1);
    });
  });

  test("measurer agrees with calculateTabWidth on the right-tab line width", () => {
    withFakeTextMeasure(() => {
      const block = {
        kind: "paragraph" as const,
        id: "right-tab-width",
        runs: [
          { kind: "text" as const, text: "Title", fontSize: 11 },
          { kind: "tab" as const },
          { kind: "text" as const, text: "7", fontSize: 11 },
        ],
        attrs: {
          tabs: [{ val: "end" as const, pos: 5000 }],
        },
      };

      const measure = measureParagraph(block, 400);
      const line = measure.lines.at(0);
      expect(line).toBeDefined();

      // What the painter computes end-to-end:
      //   title (5 chars × 5px = 25px) +
      //   tab width that anchors "7" right of the 5000-twip stop +
      //   "7" (10px)
      const titleWidth = 25;
      const followingWidth = 10;
      const tabContext: TabContext = {
        explicitStops: [{ val: "end", pos: 5000 }],
      };
      const tabResult = calculateTabWidth(titleWidth, tabContext, {
        followingWidth,
      });
      const painterLineWidth = titleWidth + tabResult.width + followingWidth;

      expect(Math.abs((line?.width ?? 0) - painterLineWidth)).toBeLessThanOrEqual(0.5);
    });
  });

  test("right tab reserves rotated inline image bbox width in following width", () => {
    withFakeTextMeasure(() => {
      const block = {
        kind: "paragraph" as const,
        id: "tab-rotated-inline",
        runs: [
          { kind: "text" as const, text: "Title", fontSize: 11 },
          { kind: "tab" as const },
          {
            kind: "image" as const,
            src: "img.png",
            width: 40,
            height: 20,
            rotation: 90,
          },
        ],
        attrs: {
          tabs: [{ val: "end" as const, pos: 5000 }],
        },
      };

      const measure = measureParagraph(block, 400);
      const line = measure.lines.at(0);
      expect(line).toBeDefined();

      const titleWidth = 25;
      const followingWidth = 20;
      const tabResult = calculateTabWidth(
        titleWidth,
        {
          explicitStops: [{ val: "end", pos: 5000 }],
        },
        {
          followingWidth,
        },
      );
      const painterLineWidth = titleWidth + tabResult.width + followingWidth;
      expect(Math.abs((line?.width ?? 0) - painterLineWidth)).toBeLessThanOrEqual(0.5);
    });
  });
});
