import { describe, expect, test } from "bun:test";

import type { PageContentBand } from "@stll/folio-core/paged-layout/rangeProjection";
import type { DirectiveRange } from "@stll/folio-core/prosemirror/plugins/templateDirectives";

import {
  closerHintLabel,
  innermostBlockAt,
  pairBlockRanges,
  railBudgetDepth,
  railXForDepth,
  segmentBandByPages,
} from "./TemplateDirectivesOverlay";

/**
 * Presentation math for the block-directive rails. The rendering itself is DOM;
 * these are the pure pieces the overlay leans on, so a regression in the gutter
 * budget, page segmentation, pairing, or caret containment fails here first.
 *
 * Constants that pin the expected numbers (must match the overlay):
 *   RAIL_WIDTH 2.5, RAIL_DEPTH_GAP 5, RAIL_TEXT_CLEARANCE 6, RAIL_EDGE_PAD 2.
 * usable = margin - 6 - 2.5 - 2 = margin - 10.5;  budget = floor(usable / 5).
 */

const blockRange = (from: number, kind: DirectiveRange["kind"], expr = ""): DirectiveRange => ({
  from,
  to: from + 1,
  kind,
  expr,
  block: true,
});

describe("railBudgetDepth", () => {
  test("no usable margin yields depth 0 (never negative)", () => {
    expect(railBudgetDepth(0)).toBe(0);
    expect(railBudgetDepth(-40)).toBe(0);
    expect(railBudgetDepth(10)).toBe(0); // usable -0.5
  });

  test("counts how many 5px steps fit after the clearances", () => {
    expect(railBudgetDepth(20)).toBe(1); // usable 9.5 -> 1
    expect(railBudgetDepth(25)).toBe(2); // usable 14.5 -> 2
    expect(railBudgetDepth(50)).toBe(7); // usable 39.5 -> 7
  });
});

describe("railXForDepth", () => {
  const CONTENT_LEFT = 100;
  const MARGIN = 50; // budget 7, pageLeft = 50

  test("deepest budgeted rail keeps a >=6px clearance to the text edge", () => {
    const x = railXForDepth(7, CONTENT_LEFT, MARGIN);
    expect(x).toBe(91.5); // 100 - 6 - 2.5
    const railRightEdge = x + 2.5; // + RAIL_WIDTH
    expect(CONTENT_LEFT - railRightEdge).toBeGreaterThanOrEqual(6);
  });

  test("outermost rail stays inside the page margin", () => {
    const x = railXForDepth(0, CONTENT_LEFT, MARGIN);
    expect(x).toBe(56.5); // 91.5 - 7*5
    const pageLeft = CONTENT_LEFT - MARGIN;
    expect(x).toBeGreaterThanOrEqual(pageLeft + 2); // + RAIL_EDGE_PAD
  });

  test("depths beyond the budget reuse the deepest offset (no march into text)", () => {
    expect(railXForDepth(99, CONTENT_LEFT, MARGIN)).toBe(railXForDepth(7, CONTENT_LEFT, MARGIN));
  });

  test("deeper depth sits closer to the text than a shallower one", () => {
    expect(railXForDepth(3, CONTENT_LEFT, MARGIN)).toBeGreaterThan(
      railXForDepth(1, CONTENT_LEFT, MARGIN),
    );
  });
});

describe("segmentBandByPages", () => {
  // Two pages with a 20px inter-page gap (page 0 body 0..100, page 1 body 120..220).
  const pages: PageContentBand[] = [
    { pageIndex: 0, top: 0, bottom: 100 },
    { pageIndex: 1, top: 120, bottom: 220 },
  ];

  test("a single-page band is clipped to that page", () => {
    expect(segmentBandByPages({ top: 10, bottom: 60 }, pages)).toEqual([{ top: 10, bottom: 60 }]);
  });

  test("a cross-page band splits and never paints over the inter-page gap", () => {
    expect(segmentBandByPages({ top: 50, bottom: 180 }, pages)).toEqual([
      { top: 50, bottom: 100 },
      { top: 120, bottom: 180 },
    ]);
  });

  test("a band that lands entirely in the gap/footer/header draws nothing", () => {
    expect(segmentBandByPages({ top: 104, bottom: 116 }, pages)).toEqual([]);
  });

  test("sub-2px slivers at a seam are dropped", () => {
    expect(segmentBandByPages({ top: 99, bottom: 100 }, pages)).toEqual([]);
  });

  test("with no page geometry the whole span is one segment (fallback)", () => {
    expect(segmentBandByPages({ top: 10, bottom: 60 }, [])).toEqual([{ top: 10, bottom: 60 }]);
  });
});

describe("pairBlockRanges", () => {
  test("pairs nested openers with their closers and carries id/kind/expr", () => {
    // {{#each contracts.risks}} > {{#if hasVerdicts}} ... {{/if}} {{/each}}
    const ranges = [
      blockRange(0, "each", "contracts.risks"),
      blockRange(10, "if", "hasVerdicts"),
      blockRange(20, "endif"),
      blockRange(30, "endeach"),
    ];

    const pairings = pairBlockRanges(ranges);

    expect(pairings).toHaveLength(2);
    const inner = pairings.find((p) => p.kind === "if");
    const outer = pairings.find((p) => p.kind === "each");
    expect(inner).toMatchObject({ blockId: 10, openerFrom: 10, closerFrom: 20, closerTo: 21 });
    expect(inner?.openerExpr).toBe("hasVerdicts");
    expect(outer).toMatchObject({ blockId: 0, closerFrom: 30, openerExpr: "contracts.risks" });
  });

  test("drops inline (block:false) markers and unbalanced closers", () => {
    const ranges: DirectiveRange[] = [
      blockRange(0, "endif"), // stray closer, no opener
      { from: 5, to: 6, kind: "if", expr: "inline", block: false },
      blockRange(10, "if", "kept"),
      blockRange(20, "endif"),
    ];

    const pairings = pairBlockRanges(ranges);

    expect(pairings).toHaveLength(1);
    expect(pairings[0]).toMatchObject({ blockId: 10, openerExpr: "kept" });
  });

  test("carries the nesting depth from the same walk", () => {
    // {{#each}} > {{#if}} ... {{/if}} {{/each}}: depth rides on the pairing so the
    // rail's indentation and its opener→closer span can never diverge.
    const pairings = pairBlockRanges([
      blockRange(0, "each", "items"),
      blockRange(10, "if", "cond"),
      blockRange(20, "endif"),
      blockRange(30, "endeach"),
    ]);

    expect(pairings.find((p) => p.kind === "if")?.depth).toBe(1);
    expect(pairings.find((p) => p.kind === "each")?.depth).toBe(0);
  });

  test("kind-aware: a {{/if}} pairs with its {{#if}}, not an intervening {{#each}}", () => {
    // {{#if}} {{#each}} {{/if}} {{/each}} (crossed nesting). Blind popping would
    // pair the {{/if}} with the {{#each}} (an each-kind band) and the {{/each}}
    // with the {{#if}}. Kind-aware matching closes the if correctly and drops the
    // improperly-nested each rather than mislabelling a band.
    const pairings = pairBlockRanges([
      blockRange(0, "if", "A"),
      blockRange(10, "each", "B"),
      blockRange(20, "endif"),
      blockRange(30, "endeach"),
    ]);

    expect(pairings).toHaveLength(1);
    expect(pairings[0]).toMatchObject({
      blockId: 0,
      kind: "if",
      openerFrom: 0,
      closerFrom: 20,
      openerExpr: "A",
    });
  });

  test("kind-aware: a {{/if}} never closes an {{#each}}", () => {
    // {{#each}} {{/if}}: mismatched families must not pair (would be an each-kind
    // band closed by a /if). Blind popping paired them; kind-aware drops both.
    const pairings = pairBlockRanges([blockRange(0, "each", "B"), blockRange(10, "endif")]);

    expect(pairings).toHaveLength(0);
  });
});

describe("closerHintLabel", () => {
  test("names what the closer closes", () => {
    expect(closerHintLabel("if", "hasVerdicts")).toBe("/if · hasVerdicts");
    expect(closerHintLabel("each", "contracts.risks")).toBe("/each · contracts.risks");
  });

  test("falls back to the bare keyword when there is no expression", () => {
    expect(closerHintLabel("if", "")).toBe("/if");
    expect(closerHintLabel("each", "")).toBe("/each");
  });
});

describe("innermostBlockAt", () => {
  const pairings = pairBlockRanges([
    blockRange(0, "each", "items"),
    blockRange(10, "if", "cond"),
    blockRange(20, "endif"),
    blockRange(30, "endeach"),
  ]);

  test("picks the deepest block whose span contains the caret", () => {
    expect(innermostBlockAt(pairings, 15)).toBe(10); // inside inner if
  });

  test("falls back to the enclosing block when outside the inner one", () => {
    expect(innermostBlockAt(pairings, 25)).toBe(0); // between /if and /each
  });

  test("returns null outside every block or with no caret", () => {
    expect(innermostBlockAt(pairings, 100)).toBeNull();
    expect(innermostBlockAt(pairings, null)).toBeNull();
  });
});
