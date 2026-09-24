/**
 * A `w:br w:type="page"` ends its page. What precedes it in the paragraph
 * stays on that page, and the page it opens is still fresh for a following
 * `w:pageBreakBefore` paragraph as long as nothing visible has been placed on
 * it yet.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { layoutDocument } from "../../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type { FlowBlock, Layout, LayoutOptions, PageMargins } from "../../layout-engine/types";
import { schema } from "../../prosemirror/schema";
import { toFlowBlocks } from "./toFlowBlocks";

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;
const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };
/** 1 inch below the anchoring paragraph. */
const BOX_OFFSET_EMU = 914_400;
const BOX_OFFSET_PX = 96;

const text = (value: string, attrs: Record<string, unknown> | null = null) =>
  schema.node("paragraph", attrs, [schema.text(value)]);
const heading = (value: string) => text(value, { pageBreakBefore: true });
const breakOnly = () => schema.node("paragraph", null, [schema.node("pageBreakRun")]);
const anchoredBox = (anchorId: string, value: string) =>
  schema.node(
    "textBox",
    {
      width: 300,
      height: 60,
      _docxAnchorId: anchorId,
      position: { vertical: { relativeTo: "paragraph", posOffset: BOX_OFFSET_EMU } },
    },
    [text(value)],
  );
const anchor = (anchorId: string) => schema.node("textBoxAnchor", { anchorId });

type LaidOut = { blocks: FlowBlock[]; layout: Layout };

const layOut = (doc: PMNode): LaidOut => {
  let laidOut: LaidOut | undefined;
  withFakeTextMeasure(
    () => {
      const blocks = toFlowBlocks(doc);
      laidOut = {
        blocks,
        layout: layoutDocument(blocks, measureBlocks(blocks, CONTENT_WIDTH), LAYOUT_OPTIONS),
      };
    },
    { charWidth: fixedCharWidth(5) },
  );
  if (!laidOut) {
    throw new Error("Expected a layout");
  }
  return laidOut;
};

/** Page index and top of the paragraph fragment painting `value`. */
const placementOf = ({ blocks, layout }: LaidOut, value: string) => {
  const block = blocks.find(
    (candidate) =>
      candidate.kind === "paragraph" &&
      candidate.runs.some((run) => run.kind === "text" && run.text === value),
  );
  for (const [pageIndex, page] of layout.pages.entries()) {
    const fragment = page.fragments.find((candidate) => candidate.blockId === block?.id);
    if (fragment) {
      return { pageIndex, y: fragment.y };
    }
  }
  return undefined;
};

const textBoxPlacement = ({ layout }: LaidOut) => {
  for (const [pageIndex, page] of layout.pages.entries()) {
    const fragment = page.fragments.find((candidate) => candidate.kind === "textBox");
    if (fragment) {
      return { pageIndex, y: fragment.y };
    }
  }
  return undefined;
};

describe("w:pageBreakBefore on a page a page break just opened", () => {
  test("starts on that page after a break-only paragraph", () => {
    const laidOut = layOut(schema.node("doc", null, [text("cover"), breakOnly(), heading("one")]));

    expect(placementOf(laidOut, "one")).toEqual({ pageIndex: 1, y: MARGINS.top });
    expect(laidOut.layout.pages).toHaveLength(2);
  });

  test("starts on that page after a break that ends a text paragraph", () => {
    const laidOut = layOut(
      schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("cover"), schema.node("pageBreakRun")]),
        heading("one"),
      ]),
    );

    expect(placementOf(laidOut, "one")).toEqual({ pageIndex: 1, y: MARGINS.top });
    expect(laidOut.layout.pages).toHaveLength(2);
  });

  test("still breaks after an empty paragraph that takes a line", () => {
    const laidOut = layOut(
      schema.node("doc", null, [
        text("cover"),
        breakOnly(),
        schema.node("paragraph"),
        heading("one"),
      ]),
    );

    expect(placementOf(laidOut, "one")?.pageIndex).toBe(2);
  });

  test("still breaks after visible content", () => {
    const laidOut = layOut(schema.node("doc", null, [text("cover"), heading("one")]));

    expect(placementOf(laidOut, "one")?.pageIndex).toBe(1);
  });
});

describe("a text box anchored before a page break in its paragraph", () => {
  test("stays on the page the break ends", () => {
    const laidOut = layOut(
      schema.node("doc", null, [
        text("title page"),
        schema.node("paragraph", null, [anchor("box"), schema.node("pageBreakRun")]),
        anchoredBox("box", "boxed title"),
        text("next page"),
      ]),
    );

    const titlePage = placementOf(laidOut, "title page");
    const box = textBoxPlacement(laidOut);

    expect(box?.pageIndex).toBe(0);
    // Nothing of the anchoring paragraph precedes the break, so it starts
    // right below the previous paragraph.
    expect(box?.y).toBeGreaterThan((titlePage?.y ?? 0) + BOX_OFFSET_PX);
    expect(placementOf(laidOut, "next page")).toEqual({ pageIndex: 1, y: MARGINS.top });
    expect(laidOut.layout.pages).toHaveLength(2);
  });

  test("is measured from the paragraph text before the break", () => {
    const laidOut = layOut(
      schema.node("doc", null, [
        text("title page"),
        schema.node("paragraph", null, [
          schema.text("lead"),
          anchor("box"),
          schema.node("pageBreakRun"),
          schema.text("tail"),
        ]),
        anchoredBox("box", "boxed title"),
      ]),
    );

    const lead = placementOf(laidOut, "lead");
    const box = textBoxPlacement(laidOut);

    expect(lead?.pageIndex).toBe(0);
    expect(box).toEqual({ pageIndex: 0, y: (lead?.y ?? 0) + BOX_OFFSET_PX });
    expect(placementOf(laidOut, "tail")?.pageIndex).toBe(1);
  });

  test("anchored after the break, moves on with the rest of its paragraph", () => {
    const laidOut = layOut(
      schema.node("doc", null, [
        text("title page"),
        schema.node("paragraph", null, [
          schema.text("lead"),
          schema.node("pageBreakRun"),
          schema.text("tail"),
          anchor("box"),
        ]),
        anchoredBox("box", "boxed title"),
      ]),
    );

    const tail = placementOf(laidOut, "tail");

    expect(tail?.pageIndex).toBe(1);
    expect(textBoxPlacement(laidOut)).toEqual({
      pageIndex: 1,
      y: (tail?.y ?? 0) + BOX_OFFSET_PX,
    });
  });
});
