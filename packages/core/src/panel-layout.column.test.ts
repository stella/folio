import { describe, expect, test } from "bun:test";

import {
  commentsColumnStart,
  computePanelLayout,
  PANEL_METRICS,
  type PanelLayout,
  type PanelLayoutInput,
} from "./panel-layout";

const LETTER_PAGE = 816;
const LANDSCAPE_PAGE = 1056;
const WIDTHS = [480, 800, 1100, 1500] as const;
const OUTLINES = ["absent", "available", "expanded"] as const;
const COMMENTS = ["closed", "open"] as const;

type ColumnSource = "layout" | "measured-before-outline-track";

/**
 * The column start the adapters use, or, with the test-side switch, the
 * value a column measured from the page gets when it measures while the
 * scroll content is still the whole row (before the outline takes its track).
 */
const columnStartFrom = (
  source: ColumnSource,
  input: PanelLayoutInput,
  layout: PanelLayout,
): number | null => {
  if (source === "layout") return layout.commentsColumnStart;
  if (layout.comments !== "column") return null;
  // The page's end edge as it sat while the scroll content was the whole row.
  const pageEnd =
    (input.availableWidth - layout.commentsGutter - input.pageWidth) / 2 +
    input.pageWidth;
  return commentsColumnStart(pageEnd, input.availableWidth);
};

/**
 * Everything wrong with a column start, against the page as the pages
 * viewport places it: the page margin on both sides, the comments gutter
 * after the page, the page centred in what is left. Positions are in the
 * row the outline track and the scroll viewport share.
 */
const columnViolations = (
  input: PanelLayoutInput,
  layout: PanelLayout,
  start: number | null,
): string[] => {
  const { commentsGap, commentsWidth, pageMargin } = PANEL_METRICS;
  if (layout.comments !== "column") {
    return start === null ? [] : ["a column start without a comments column"];
  }
  if (start === null) return ["a comments column without a start"];
  const viewportWidth = input.availableWidth - layout.outlineTrackWidth;
  const pageStart =
    layout.outlineTrackWidth +
    pageMargin +
    (viewportWidth - 2 * pageMargin - layout.commentsGutter - input.pageWidth) /
      2;
  const pageEnd = pageStart + input.pageWidth;
  const columnLeft = layout.outlineTrackWidth + start;
  const violations: string[] = [];
  if (Math.abs(columnLeft - (pageEnd + commentsGap)) > 0.5) {
    violations.push(
      `column starts at ${String(columnLeft)}, not a gap after the page end ${String(pageEnd)}`,
    );
  }
  if (columnLeft + commentsWidth + commentsGap > input.availableWidth + 0.5) {
    violations.push(
      `column ends at ${String(columnLeft + commentsWidth)}, past the row ${String(input.availableWidth)}`,
    );
  }
  if (columnLeft < layout.outlineTrackWidth) {
    violations.push("column starts under the outline track");
  }
  return violations;
};

const cases = (pageWidth: number): PanelLayoutInput[] =>
  WIDTHS.flatMap((availableWidth) =>
    OUTLINES.flatMap((outline) =>
      COMMENTS.map((comments) => ({
        availableWidth,
        pageWidth,
        outline,
        comments,
      })),
    ),
  );

const caseName = (input: PanelLayoutInput) =>
  `${String(input.availableWidth)}px, ${String(input.pageWidth)}px page, outline ${input.outline}, comments ${input.comments}`;

describe("the comments column start the layout computes", () => {
  test("sits a gap after the page and inside the row at every tier", () => {
    for (const pageWidth of [LETTER_PAGE, LANDSCAPE_PAGE]) {
      for (const input of cases(pageWidth)) {
        const layout = computePanelLayout(input);
        expect({
          name: caseName(input),
          violations: columnViolations(
            input,
            layout,
            columnStartFrom("layout", input, layout),
          ),
        }).toEqual({ name: caseName(input), violations: [] });
      }
    }
  });

  test("places the wide tier's column beside the page the outline moved", () => {
    const columnLeft = (input: PanelLayoutInput) => {
      const layout = computePanelLayout(input);
      return layout.commentsColumnStart === null
        ? null
        : layout.outlineTrackWidth + layout.commentsColumnStart;
    };
    const wide = { availableWidth: 1500, pageWidth: LETTER_PAGE } as const;
    // The outline column takes 256 px; the page sits at 318..1134.
    expect(
      columnLeft({ ...wide, outline: "available", comments: "open" }),
    ).toBe(1134 + PANEL_METRICS.commentsGap);
    expect(columnLeft({ ...wide, outline: "expanded", comments: "open" })).toBe(
      1134 + PANEL_METRICS.commentsGap,
    );
    // No outline: the page sits at 190..1006.
    expect(columnLeft({ ...wide, outline: "absent", comments: "open" })).toBe(
      1006 + PANEL_METRICS.commentsGap,
    );
    // The widest page of mixed sections sets the room: a rail, then the page
    // at 86..1142.
    expect(
      columnLeft({
        availableWidth: 1500,
        pageWidth: LANDSCAPE_PAGE,
        outline: "available",
        comments: "open",
      }),
    ).toBe(1142 + PANEL_METRICS.commentsGap);
    for (const width of [480, 800, 1100]) {
      expect(
        columnLeft({
          availableWidth: width,
          pageWidth: LETTER_PAGE,
          outline: "available",
          comments: "open",
        }),
      ).toBeNull();
    }
  });

  test("is unknown until the row is measured", () => {
    expect(
      computePanelLayout({
        availableWidth: Number.POSITIVE_INFINITY,
        pageWidth: LETTER_PAGE,
        outline: "available",
        comments: "open",
      }).commentsColumnStart,
    ).toBeNull();
  });

  test("a start measured before the outline took its track is caught", () => {
    const input = {
      availableWidth: 1500,
      pageWidth: LETTER_PAGE,
      outline: "available",
      comments: "open",
    } satisfies PanelLayoutInput;
    const layout = computePanelLayout(input);
    const stale = columnStartFrom(
      "measured-before-outline-track",
      input,
      layout,
    );
    expect(stale).toBe(1006 + PANEL_METRICS.commentsGap);
    expect(columnViolations(input, layout, stale).length).toBeGreaterThan(0);
  });
});

describe("commentsColumnStart", () => {
  test("starts a gap after the page and stays on its surface", () => {
    const { commentsGap, commentsWidth } = PANEL_METRICS;
    // Room beside the page: the column follows the page's end edge.
    expect(commentsColumnStart(900, 1244)).toBe(900 + commentsGap);
    // A page ending near the surface's edge: the column stays on it.
    expect(commentsColumnStart(1100, 1244)).toBe(1244 - commentsWidth - 8);
    // A surface narrower than the column: it keeps its start inset.
    expect(commentsColumnStart(100, 200)).toBe(8);
  });
});
