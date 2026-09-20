/**
 * What a bidirectional wrapper does to the page.
 *
 * The mark says the run was authored inside `w:bdo` or `w:dir`; the painter
 * has to draw that, or the carrier is a record nothing reads. The assertion is
 * on the flow block and the glyph run rather than on pixels: the direction a
 * run is laid out in is a fact of the layout, and a screenshot would be a
 * slower way of asking the same question.
 *
 * Before the paint leg existed the flow run carried no `bidiWrapper` and the
 * glyph run took the paragraph's direction, so both cases below failed.
 */

import { describe, expect, test } from "bun:test";

import { buildDisplayList } from "../../display-list/build/buildDisplayList";
import type { DisplayGlyphRun } from "../../display-list/types";
import { layoutDocument } from "../../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type { LayoutOptions, PageMargins } from "../../layout-engine/types";
import { schema } from "../../prosemirror/schema";
import type { InlineWrapperLayer } from "../../prosemirror/schema/marks";
import { toFlowBlocks } from "./toFlowBlocks";

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;
const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };

const wrapped = (stack: readonly InlineWrapperLayer[]) =>
  schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("abc", [schema.marks["inlineWrapper"]!.create({ stack })]),
    ]),
  ]);

const firstTextRun = (stack: readonly InlineWrapperLayer[]) => {
  const block = toFlowBlocks(wrapped(stack)).at(0);
  if (block?.kind !== "paragraph") {
    throw new Error("Expected a paragraph block");
  }
  return block.runs.find((run) => run.kind === "text");
};

describe("a run the author wrapped in a bidirectional container", () => {
  test("an override reaches the layout engine as an override", () => {
    expect(
      firstTextRun([{ kind: "bidi", control: "override", direction: "rtl" }])?.bidiWrapper,
    ).toEqual({ control: "override", direction: "rtl" });
  });

  test("an embedding reaches it as an embedding", () => {
    expect(
      firstTextRun([{ kind: "bidi", control: "embedding", direction: "ltr" }])?.bidiWrapper,
    ).toEqual({ control: "embedding", direction: "ltr" });
  });

  test("a wrapper that states no direction states none here either", () => {
    expect(firstTextRun([{ kind: "bidi", control: "embedding" }])?.bidiWrapper).toEqual({
      control: "embedding",
    });
  });

  test("the innermost wrapper is the one the run is laid out in", () => {
    expect(
      firstTextRun([
        { kind: "bidi", control: "embedding", direction: "ltr" },
        { kind: "bidi", control: "override", direction: "rtl" },
      ])?.bidiWrapper,
    ).toEqual({ control: "override", direction: "rtl" });
  });

  test("a run with no wrapper states none", () => {
    const block = toFlowBlocks(
      schema.node("doc", null, [schema.node("paragraph", null, [schema.text("abc")])]),
    ).at(0);
    if (block?.kind !== "paragraph") {
      throw new Error("Expected a paragraph block");
    }
    expect(block.runs.find((run) => run.kind === "text")?.bidiWrapper).toBeUndefined();
  });
});

/**
 * The direction a glyph run is painted in is where the wrapper stops being a
 * record and becomes layout. Every backend — the DOM painter, the PDF writer —
 * lays out by this one field, so asserting on it is asserting on the page.
 */
describe("the glyph run a wrapped run paints", () => {
  const paintedDirections = (stack: readonly InlineWrapperLayer[]): string[] => {
    let directions: string[] = [];
    withFakeTextMeasure(
      () => {
        const blocks = toFlowBlocks(wrapped(stack));
        const measures = measureBlocks(blocks, CONTENT_WIDTH);
        // Derived from the builder rather than imported: the painter owns the
        // type, and layout-bridge may not reach into that layer.
        const blockLookup: Parameters<typeof buildDisplayList>[0]["blockLookup"] = new Map();
        for (const [index, block] of blocks.entries()) {
          const measure = measures[index];
          if (measure) {
            blockLookup.set(String(block.id), { block, measure });
          }
        }
        const list = buildDisplayList({
          layout: layoutDocument(blocks, measures, LAYOUT_OPTIONS),
          blockLookup,
        });
        directions = (list.pages.at(0)?.primitives ?? [])
          .filter((primitive): primitive is DisplayGlyphRun => primitive.kind === "glyphRun")
          .map((run) => run.direction);
      },
      { charWidth: fixedCharWidth(5) },
    );
    return directions;
  };

  test("an rtl override paints right to left inside a left-to-right paragraph", () => {
    expect(paintedDirections([{ kind: "bidi", control: "override", direction: "rtl" }])).toEqual([
      "rtl",
    ]);
  });

  test("a wrapper that states no direction leaves the paragraph's", () => {
    expect(paintedDirections([{ kind: "bidi", control: "override" }])).toEqual(["ltr"]);
  });
});
