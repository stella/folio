import { describe, expect, test } from "bun:test";

import {
  INITIAL_RENDERED_BREAK_STATE,
  reconcileAfterBlock,
  reconcileBreakBeforeBlock,
  recordReflowBoundary,
} from "./renderedBreakReconciliation";
import type { FlowBlock, Page, ParagraphBlock } from "./types";

const paragraph = (
  id: number,
  attrs: ParagraphBlock["attrs"] = {},
  runs: ParagraphBlock["runs"] = [{ kind: "text", text: `Paragraph ${id}` }],
): ParagraphBlock => ({ kind: "paragraph", id, attrs, runs });

const page = (fragments: Page["fragments"] = []): Page => ({
  number: 1,
  fragments,
  margins: { top: 10, right: 10, bottom: 10, left: 10 },
  size: { w: 100, h: 100 },
});

const paragraphFragment = (
  blockId: number,
  continuesFromPrev = false,
): Page["fragments"][number] => ({
  kind: "paragraph",
  blockId,
  x: 0,
  y: 0,
  width: 80,
  height: 10,
  fromLine: 0,
  toLine: 1,
  ...(continuesFromPrev ? { continuesFromPrev: true } : {}),
});

describe("rendered break reconciliation", () => {
  test("explicit breaks force a page and consume prior advance state", () => {
    const block = paragraph(2);
    const decision = reconcileBreakBeforeBlock({
      state: { type: "pageAdvance", reason: "ordinary" },
      block,
      previousBlock: paragraph(1),
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", paragraph(1)]]),
      hasExplicitPageBreak: true,
      renderedBreakNeedsSnap: false,
    });

    expect(decision).toEqual({
      forcePageBreak: true,
      suppressSpaceBefore: false,
      state: INITIAL_RENDERED_BREAK_STATE,
    });
  });

  test("an unsatisfied cached marker forces a page after visible content", () => {
    const previous = paragraph(1);
    const marker = paragraph(2, { renderedPageBreakBefore: true });
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: marker,
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(true);
    expect(decision.suppressSpaceBefore).toBe(true);
    expect(decision.state).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("a fitting cached marker remains advisory", () => {
    const previous = paragraph(1);
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: paragraph(2, { renderedPageBreakBefore: true }),
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.suppressSpaceBefore).toBe(false);
    expect(decision.state).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("a fitting cached marker remains authoritative on a keep-next paragraph", () => {
    const previous = paragraph(1);
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: paragraph(2, { keepNext: true, renderedPageBreakBefore: true }),
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.forcePageBreak).toBe(true);
    expect(decision.state).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("paragraph continuation satisfies a cached marker", () => {
    const previous = paragraph(1);
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: paragraph(2, { renderedPageBreakBefore: true }),
      previousBlock: previous,
      page: page([paragraphFragment(1, true)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.suppressSpaceBefore).toBe(true);
  });

  test("a section boundary preserves explicit paragraph spacing", () => {
    const previous: FlowBlock = { kind: "sectionBreak", id: 1, type: "nextPage" };
    const marker = paragraph(2, {
      renderedPageBreakBefore: true,
      spacing: { before: 24 },
      spacingExplicit: { before: true },
    });
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: marker,
      previousBlock: previous,
      page: page(),
      blocksById: new Map(),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.suppressSpaceBefore).toBe(false);
  });

  test("a section boundary consumes inherited paragraph spacing", () => {
    const previous: FlowBlock = { kind: "sectionBreak", id: 1, type: "nextPage" };
    const marker = paragraph(2, {
      renderedPageBreakBefore: true,
      spacing: { before: 24 },
    });
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: marker,
      previousBlock: previous,
      page: page(),
      blocksById: new Map(),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.suppressSpaceBefore).toBe(true);
  });

  test("a continuous section boundary preserves spacing when the marker snaps", () => {
    const prior = paragraph(1);
    const previous: FlowBlock = { kind: "sectionBreak", id: 2, type: "continuous" };
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: paragraph(3, {
        renderedPageBreakBefore: true,
        spacingExplicit: { before: true },
      }),
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", prior]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(true);
    expect(decision.suppressSpaceBefore).toBe(false);
  });

  test("a continuous section boundary preserves inherited spacing without a page advance", () => {
    const prior = paragraph(1);
    const previous: FlowBlock = { kind: "sectionBreak", id: 2, type: "continuous" };
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: paragraph(3, {
        renderedPageBreakBefore: true,
        spacing: { before: 24 },
      }),
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", prior]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.suppressSpaceBefore).toBe(false);
  });

  test("an authored page boundary remains satisfied across an empty carrier", () => {
    const empty = paragraph(2, {}, []);
    const boundaryState = reconcileAfterBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: { kind: "pageBreak", id: 1 },
      pageNumberBefore: 1,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const carriedState = reconcileAfterBlock({
      state: boundaryState,
      block: empty,
      pageNumberBefore: 2,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const decision = reconcileBreakBeforeBlock({
      state: carriedState,
      block: paragraph(3, {
        renderedPageBreakBefore: true,
        spacing: { before: 24 },
      }),
      previousBlock: empty,
      page: page([paragraphFragment(2)]),
      blocksById: new Map([["2", empty]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.suppressSpaceBefore).toBe(true);
    expect(decision.state).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("an authored section boundary preserves explicit spacing across an empty carrier", () => {
    const empty = paragraph(2, {}, []);
    const boundaryState = reconcileAfterBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: { kind: "sectionBreak", id: 1, type: "nextPage" },
      pageNumberBefore: 1,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const decision = reconcileBreakBeforeBlock({
      state: reconcileAfterBlock({
        state: boundaryState,
        block: empty,
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
      block: paragraph(3, {
        renderedPageBreakBefore: true,
        spacing: { before: 24 },
        spacingExplicit: { before: true },
      }),
      previousBlock: empty,
      page: page([paragraphFragment(2)]),
      blocksById: new Map([["2", empty]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(decision.suppressSpaceBefore).toBe(false);
  });

  test("a coalesced section boundary upgrades carried section-spacing semantics", () => {
    const empty = paragraph(2, {}, []);
    const pageBreakState = reconcileAfterBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: { kind: "pageBreak", id: 1 },
      pageNumberBefore: 1,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const sectionState = reconcileAfterBlock({
      state: pageBreakState,
      block: { kind: "sectionBreak", id: 2, type: "continuous" },
      pageNumberBefore: 2,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const decision = reconcileBreakBeforeBlock({
      state: reconcileAfterBlock({
        state: sectionState,
        block: empty,
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
      block: paragraph(3, {
        renderedPageBreakBefore: true,
        spacing: { before: 24 },
        spacingExplicit: { before: true },
      }),
      previousBlock: empty,
      page: page([paragraphFragment(2)]),
      blocksById: new Map([["2", empty]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(sectionState).toEqual({
      type: "pageAdvance",
      reason: "authoredBoundary",
      boundary: "sectionBreak",
    });
    expect(decision.suppressSpaceBefore).toBe(false);
  });

  test("a same-page column boundary preserves carried hard-break semantics", () => {
    const empty = paragraph(3, {}, []);
    const pageBreakState = reconcileAfterBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: { kind: "pageBreak", id: 1 },
      pageNumberBefore: 1,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const columnState = reconcileAfterBlock({
      state: pageBreakState,
      block: { kind: "columnBreak", id: 2 },
      pageNumberBefore: 2,
      pageNumberAfter: 2,
      previousPage: page(),
    });
    const decision = reconcileBreakBeforeBlock({
      state: reconcileAfterBlock({
        state: columnState,
        block: empty,
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
      block: paragraph(4, {
        renderedPageBreakBefore: true,
        spacing: { before: 24 },
      }),
      previousBlock: empty,
      page: page([paragraphFragment(3)]),
      blocksById: new Map([["3", empty]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: false,
    });

    expect(columnState).toEqual({
      type: "pageAdvance",
      reason: "authoredBoundary",
      boundary: "hardBreak",
    });
    expect(decision.suppressSpaceBefore).toBe(true);
  });

  test("a reflow boundary satisfies the next cached marker", () => {
    const previous = paragraph(1);
    const decision = reconcileBreakBeforeBlock({
      state: recordReflowBoundary(INITIAL_RENDERED_BREAK_STATE, true),
      block: paragraph(2, { renderedPageBreakBefore: true }),
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(false);
    expect(decision.state).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("ordinary movement satisfies a marker within one numbered sequence", () => {
    const previous = paragraph(1, { numPr: { numId: 4, ilvl: 1 } });
    const marker = paragraph(2, {
      numPr: { numId: 4, ilvl: 1 },
      renderedPageBreakBefore: true,
    });
    const decision = reconcileBreakBeforeBlock({
      state: { type: "pageAdvance", reason: "ordinary" },
      block: marker,
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(false);
  });

  test("different tab layouts leave a cached marker authoritative", () => {
    const previous = paragraph(1, { styleId: "Tabular", tabs: [{ val: "start", pos: 720 }] }, [
      { kind: "text", text: "First" },
      { kind: "tab" },
    ]);
    const marker = paragraph(
      2,
      {
        styleId: "Tabular",
        tabs: [{ val: "start", pos: 1440 }],
        renderedPageBreakBefore: true,
      },
      [{ kind: "text", text: "Second" }, { kind: "tab" }],
    );
    const decision = reconcileBreakBeforeBlock({
      state: { type: "pageAdvance", reason: "ordinary" },
      block: marker,
      previousBlock: previous,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", previous]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(true);
  });

  test("pages containing only empty paragraphs do not force another break", () => {
    const empty = paragraph(1, {}, []);
    const decision = reconcileBreakBeforeBlock({
      state: INITIAL_RENDERED_BREAK_STATE,
      block: paragraph(2, { renderedPageBreakBefore: true }),
      previousBlock: empty,
      page: page([paragraphFragment(1)]),
      blocksById: new Map([["1", empty]]),
      hasExplicitPageBreak: false,
      renderedBreakNeedsSnap: true,
    });

    expect(decision.forcePageBreak).toBe(false);
  });
});

describe("rendered break state transitions", () => {
  test("whole-block page movement records an ordinary advance", () => {
    const block = paragraph(2);
    expect(
      reconcileAfterBlock({
        state: INITIAL_RENDERED_BREAK_STATE,
        block,
        pageNumberBefore: 1,
        pageNumberAfter: 2,
        previousPage: page([paragraphFragment(1)]),
      }),
    ).toEqual({ type: "pageAdvance", reason: "ordinary" });
  });

  test("split-block page movement records a reflow boundary", () => {
    const block = paragraph(2);
    expect(
      reconcileAfterBlock({
        state: INITIAL_RENDERED_BREAK_STATE,
        block,
        pageNumberBefore: 1,
        pageNumberAfter: 2,
        previousPage: page([paragraphFragment(2)]),
      }),
    ).toEqual({ type: "pageAdvance", reason: "reflowBoundary" });
  });

  test("page-breaking structural blocks record an authored boundary", () => {
    const block: FlowBlock = { kind: "pageBreak", id: 2 };
    expect(
      reconcileAfterBlock({
        state: { type: "pageAdvance", reason: "ordinary" },
        block,
        pageNumberBefore: 1,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
    ).toEqual({
      type: "pageAdvance",
      reason: "authoredBoundary",
      boundary: "hardBreak",
    });
  });

  test("a visible block consumes an authored boundary before a later marker", () => {
    expect(
      reconcileAfterBlock({
        state: {
          type: "pageAdvance",
          reason: "authoredBoundary",
          boundary: "hardBreak",
        },
        block: paragraph(2),
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
    ).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("a visible list marker consumes an authored boundary on a run-empty paragraph", () => {
    expect(
      reconcileAfterBlock({
        state: {
          type: "pageAdvance",
          reason: "authoredBoundary",
          boundary: "hardBreak",
        },
        block: paragraph(2, { listMarker: "1." }, []),
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
    ).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("visible shading consumes an authored boundary on a run-empty paragraph", () => {
    expect(
      reconcileAfterBlock({
        state: {
          type: "pageAdvance",
          reason: "authoredBoundary",
          boundary: "hardBreak",
        },
        block: paragraph(2, { shading: "#FFFF00" }, []),
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
    ).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("a visible border consumes an authored boundary on a run-empty paragraph", () => {
    expect(
      reconcileAfterBlock({
        state: {
          type: "pageAdvance",
          reason: "authoredBoundary",
          boundary: "hardBreak",
        },
        block: paragraph(
          2,
          { borders: { bottom: { style: "solid", width: 1, color: "#000000" } } },
          [],
        ),
        pageNumberBefore: 2,
        pageNumberAfter: 2,
        previousPage: page(),
      }),
    ).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });

  test("a non-page-breaking structural block clears accumulated movement", () => {
    expect(
      reconcileAfterBlock({
        state: { type: "pageAdvance", reason: "ordinary" },
        block: { kind: "sectionBreak", id: 2, type: "continuous" },
        pageNumberBefore: 1,
        pageNumberAfter: 1,
        previousPage: page(),
      }),
    ).toEqual(INITIAL_RENDERED_BREAK_STATE);
  });
});
