/**
 * A paragraph holding only `w:br w:type="page"` keeps its paragraph mark on
 * the page the break ends unless `w:splitPgBreakAndParaMark` is set. The flow
 * projection still maps that mark after the break, so the carrier must not
 * bring the paragraph's before/after spacing onto the next page.
 */

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "../../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type { FlowBlock, Layout, LayoutOptions, PageMargins } from "../../layout-engine/types";
import type { Node as PMNode } from "prosemirror-model";
import { schema } from "../../prosemirror/schema";
import { toFlowBlocks } from "./toFlowBlocks";

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;
const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };

// Directly authored spacing survives the empty-paragraph spacing collapse.
const spaced = {
  spaceBefore: 120,
  spaceAfter: 240,
  spacingExplicit: { before: true, after: true },
};

const breakOnlyParagraph = (attrs: Record<string, unknown> = {}) =>
  schema.node("paragraph", { ...spaced, ...attrs }, [schema.node("pageBreakRun")]);

type LaidOut = { result: Layout; lastBlock: FlowBlock };

const layout = (doc: PMNode): LaidOut => {
  let laidOut: LaidOut | undefined;
  withFakeTextMeasure(
    () => {
      const blocks = toFlowBlocks(doc);
      const lastBlock = blocks.at(-1);
      if (!lastBlock) {
        throw new Error("Expected flow blocks");
      }
      laidOut = {
        result: layoutDocument(blocks, measureBlocks(blocks, CONTENT_WIDTH), LAYOUT_OPTIONS),
        lastBlock,
      };
    },
    { charWidth: fixedCharWidth(5) },
  );
  if (!laidOut) {
    throw new Error("Expected a layout");
  }
  return laidOut;
};

const fragmentOf = ({ result, lastBlock }: LaidOut) => {
  for (const [pageIndex, page] of result.pages.entries()) {
    const fragment = page.fragments.find((candidate) => candidate.blockId === lastBlock.id);
    if (fragment) {
      return { pageIndex, y: fragment.y };
    }
  }
  return undefined;
};

describe("a break-only paragraph mark that stays with its page break", () => {
  test("a following nextPage section reuses the page the break opened", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First page")]),
      breakOnlyParagraph(),
      schema.node("paragraph", { _sectionProperties: {} }),
      schema.node("paragraph", null, [schema.text("Next section")]),
    ]);

    const laidOut = layout(doc);

    expect(laidOut.result.pages).toHaveLength(2);
    expect(fragmentOf(laidOut)?.pageIndex).toBe(1);
  });

  test("the paragraph after the break starts at the top margin", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First page")]),
      breakOnlyParagraph(),
      schema.node("paragraph", null, [schema.text("Second page")]),
    ]);

    const laidOut = layout(doc);

    expect(laidOut.result.pages).toHaveLength(2);
    expect(fragmentOf(laidOut)).toEqual({ pageIndex: 1, y: MARGINS.top });
  });

  test("a split paragraph mark keeps its spacing on the next page", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First page")]),
      breakOnlyParagraph(),
      schema.node("paragraph", null, [schema.text("Second page")]),
    ]);

    const carrier = toFlowBlocks(doc, { splitPageBreakAndParagraphMark: true }).find(
      (block) => block.kind === "paragraph" && block.runs.length === 0,
    );

    expect(carrier?.kind === "paragraph" && carrier.attrs?.spacing?.after).toBeGreaterThan(0);
  });
});
