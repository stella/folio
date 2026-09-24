/**
 * An empty paragraph whose only role is to carry `w:pPr/w:sectPr` ends its
 * section without a line of its own. These tests pin down where that marker
 * still matters: a page break in the paragraph before it, and markers that
 * paint a number or border or that make up their whole section.
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

const text = (value: string, attrs: Record<string, unknown> | null = null) =>
  schema.node("paragraph", attrs, [schema.text(value)]);
const sectionMark = (attrs: Record<string, unknown> = {}) =>
  schema.node("paragraph", { _sectionProperties: {}, ...attrs });

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

/** Page index and top of the fragment painting `value`. */
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

describe("a page break before an empty section mark", () => {
  const continuousDoc = (breakParagraphs: PMNode[]) =>
    schema.node("doc", { _finalSectionStart: "continuous" }, [
      ...breakParagraphs,
      sectionMark(),
      text("two"),
    ]);

  test("still starts the next page when the break ends the text paragraph", () => {
    const laidOut = layOut(
      continuousDoc([
        schema.node("paragraph", null, [schema.text("one"), schema.node("pageBreakRun")]),
      ]),
    );

    expect(placementOf(laidOut, "one")?.pageIndex).toBe(0);
    expect(placementOf(laidOut, "two")).toEqual({ pageIndex: 1, y: MARGINS.top });
    expect(laidOut.layout.pages).toHaveLength(2);
  });

  test("still starts the next page when the break has a paragraph of its own", () => {
    const laidOut = layOut(
      continuousDoc([text("one"), schema.node("paragraph", null, [schema.node("pageBreakRun")])]),
    );

    expect(placementOf(laidOut, "two")?.pageIndex).toBe(1);
    expect(laidOut.layout.pages).toHaveLength(2);
  });

  test("advances only one page before a nextPage section", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("one"), schema.node("pageBreakRun")]),
      sectionMark(),
      text("two"),
    ]);

    const laidOut = layOut(doc);

    expect(placementOf(laidOut, "two")?.pageIndex).toBe(1);
    expect(laidOut.layout.pages).toHaveLength(2);
  });

  test("still coalesces a break that ends the section-ending paragraph itself", () => {
    const doc = schema.node("doc", { _finalSectionStart: "continuous" }, [
      schema.node("paragraph", { _sectionProperties: {} }, [
        schema.text("one"),
        schema.node("pageBreakRun"),
      ]),
      text("two"),
    ]);

    expect(placementOf(layOut(doc), "two")?.pageIndex).toBe(0);
  });
});

describe("an empty section mark that paints or makes up its section", () => {
  const numbered = {
    numPr: { kind: "reference", numId: 1, ilvl: 0 },
    listMarker: "1.",
  };

  test("keeps a numbered mark after content in place despite w:pageBreakBefore", () => {
    const doc = schema.node("doc", { _finalSectionStart: "continuous" }, [
      text("one"),
      sectionMark({ ...numbered, pageBreakBefore: true }),
      text("two"),
    ]);

    const laidOut = layOut(doc);
    const mark = laidOut.blocks.find(
      (block) => block.kind === "paragraph" && block.runs.length === 0,
    );

    expect(mark?.kind === "paragraph" && mark.attrs?.pageBreakBefore).toBeFalsy();
    expect(placementOf(laidOut, "two")?.pageIndex).toBe(0);
    expect(laidOut.layout.pages).toHaveLength(1);
  });

  test("paints a bordered mark", () => {
    const doc = schema.node("doc", { _finalSectionStart: "continuous" }, [
      text("one"),
      sectionMark({ borders: { bottom: { style: "single", size: 8, color: { rgb: "000000" } } } }),
      text("two"),
    ]);

    const mark = toFlowBlocks(doc).at(1);

    expect(mark?.kind).toBe("paragraph");
    expect(mark?.kind === "paragraph" && mark.attrs?.borders?.bottom).toBeTruthy();
  });

  test("still drops a mark whose borders are all none", () => {
    const doc = schema.node("doc", { _finalSectionStart: "continuous" }, [
      text("one"),
      sectionMark({ borders: { bottom: { style: "none" } } }),
      text("two"),
    ]);

    expect(toFlowBlocks(doc).map((block) => block.kind)).toEqual([
      "paragraph",
      "sectionBreak",
      "paragraph",
    ]);
  });

  test("keeps the line of a continuous section made only of its mark", () => {
    const doc = schema.node("doc", { _finalSectionStart: "continuous" }, [
      text("one", { _sectionProperties: {} }),
      sectionMark({ _sectionProperties: { sectionStart: "continuous" } }),
      text("two"),
    ]);
    const withoutMarkOnlySection = schema.node("doc", { _finalSectionStart: "continuous" }, [
      text("one", { _sectionProperties: {} }),
      text("two"),
    ]);

    const markOnly = placementOf(layOut(doc), "two");
    const direct = placementOf(layOut(withoutMarkOnlySection), "two");

    expect(markOnly?.pageIndex).toBe(0);
    expect(direct?.pageIndex).toBe(0);
    expect(markOnly?.y ?? 0).toBeGreaterThan(direct?.y ?? 0);
  });

  test("gives a nextPage section made only of its mark its own page", () => {
    const doc = schema.node("doc", null, [
      text("one", { _sectionProperties: {} }),
      sectionMark(),
      text("two"),
    ]);

    const laidOut = layOut(doc);

    expect(placementOf(laidOut, "two")?.pageIndex).toBe(2);
    expect(laidOut.layout.pages).toHaveLength(3);
  });

  test("honours w:pageBreakBefore on a mark that makes up its section", () => {
    const doc = schema.node("doc", { _finalSectionStart: "continuous" }, [
      text("one", { _sectionProperties: {} }),
      sectionMark({ _sectionProperties: { sectionStart: "continuous" }, pageBreakBefore: true }),
      text("two"),
    ]);

    expect(placementOf(layOut(doc), "two")?.pageIndex).toBe(1);
  });
});
