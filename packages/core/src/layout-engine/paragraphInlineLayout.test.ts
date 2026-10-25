import { describe, expect, test } from "bun:test";

import {
  columnPictureClipViewportForFragment,
  pageHorizontalColumns,
  paragraphInlinePictureViewport,
  resolveParagraphLineGeometry,
  resolvePhysicalParagraphInlineLayout,
} from "../utils/paragraphInlineLayout";
import { setParagraphFrame } from "./paragraphFrame";
import type { ParagraphAttrs, ParagraphBlock } from "./types";

const resolve = (attrs?: ParagraphAttrs, text = "text") =>
  resolvePhysicalParagraphInlineLayout({
    kind: "paragraph",
    id: "paragraph",
    runs: [{ kind: "text", text, rtl: text === "مرحبا" }],
    attrs,
  });

describe("physical paragraph inline layout", () => {
  test("mirrors explicit horizontal alignment and asymmetric indents under bidi", () => {
    expect(
      resolve({
        alignment: "right",
        bidi: true,
        indent: { left: 12, right: 34 },
      }),
    ).toMatchObject({
      alignment: "left",
      explicitAlignment: "left",
      indentLeft: 34,
      indentRight: 12,
      isRtl: true,
    });
    expect(resolve({ alignment: "left", bidi: true })).toMatchObject({
      alignment: "right",
      explicitAlignment: "right",
    });
  });

  test("preserves center and justify under bidi", () => {
    expect(resolve({ alignment: "center", bidi: true }).alignment).toBe(
      "center",
    );
    expect(resolve({ alignment: "justify", bidi: true }).alignment).toBe(
      "justify",
    );
  });

  test("does not mirror explicit false or inferred RTL paragraphs", () => {
    expect(
      resolve({
        alignment: "right",
        bidi: false,
        indent: { left: 12, right: 34 },
      }),
    ).toMatchObject({
      alignment: "right",
      indentLeft: 12,
      indentRight: 34,
      isRtl: false,
    });
    expect(
      resolve({ alignment: "left", indent: { left: 12, right: 34 } }, "مرحبا"),
    ).toMatchObject({
      alignment: "left",
      indentLeft: 12,
      indentRight: 34,
      isRtl: true,
    });
  });

  test("defaults RTL paragraphs to physical right without inventing explicit alignment", () => {
    expect(resolve({ bidi: true })).toEqual({
      alignment: "right",
      indentLeft: 0,
      indentRight: 0,
      isRtl: true,
    });
  });
});

describe("paragraph line geometry", () => {
  test("resolves image alignment from the same indented floating line origin", () => {
    const geometry = resolveParagraphLineGeometry({
      fragment: { x: 100, width: 200 },
      line: {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 1,
        width: 100,
        ascent: 14,
        descent: 4,
        lineHeight: 20,
        leftOffset: 20,
        rightOffset: 10,
      },
      lineTopYPx: 25,
      alignment: "center",
      isRtl: false,
      indentLeft: 10,
      indentRight: 5,
      firstLineOffsetPx: 15,
      markerInlineWidthPx: 5,
      isFirstLine: true,
      imageOnlyRun: {
        kind: "image",
        src: "data:image/png;base64,",
        width: 100,
        height: 20,
      },
    });

    expect(geometry.alignment).toBe("center");
    expect(geometry.layoutStartXPx).toBe(50);
    expect(geometry.paintStartXPx).toBe(167.5);
  });

  test("frames use their text area before page-column matching", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "framed-picture",
      runs: [
        {
          kind: "image",
          src: "data:image/png;base64,",
          width: 180,
          height: 40,
        },
      ],
    };
    setParagraphFrame(block, { width: 150 });
    const columns = pageHorizontalColumns({
      layout: { count: 2, gap: 20, widths: [100, 200], gaps: [30] },
      contentLeft: 10,
      contentWidth: 330,
    });

    expect(
      paragraphInlinePictureViewport({
        block,
        fragment: { x: -40, width: 150 },
        paper: { left: 0, right: 400 },
        columns,
      }),
    ).toEqual({ left: -40, right: 110 });
  });

  test("matches fragments to authored column starts, not nearby indents", () => {
    const columns = pageHorizontalColumns({
      layout: { count: 2, gap: 20, widths: [100, 200], gaps: [30] },
      contentLeft: 10,
      contentWidth: 330,
    });
    expect(
      columnPictureClipViewportForFragment({
        paper: { left: 0, right: 400 },
        columns,
        fragmentX: 10,
        fragmentWidth: 100,
      }),
    ).toEqual({ left: 0, right: 125 });
    expect(
      columnPictureClipViewportForFragment({
        paper: { left: 0, right: 400 },
        columns,
        fragmentX: 140,
        fragmentWidth: 200,
      }),
    ).toEqual({ left: 125, right: 400 });
    expect(() =>
      columnPictureClipViewportForFragment({
        paper: { left: 0, right: 400 },
        columns,
        fragmentX: 135,
        fragmentWidth: 200,
      }),
    ).toThrow("Paragraph fragment has no matching page column clip bounds");
  });
});
