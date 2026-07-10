import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import type { Layout } from "@stll/folio-core/layout-engine";
import { getPageTextFromLayout } from "./pageText";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
    },
    text: { group: "inline" },
  },
});

const doc = schema.node("doc", null, [
  schema.node("paragraph", null, [schema.text("Page one first line")]),
  schema.node("paragraph", null, [schema.text("Page one second line")]),
  schema.node("paragraph", null, [schema.text("Page two content")]),
]);

const layout = {
  pages: [
    {
      number: 1,
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
      size: { w: 816, h: 1056 },
      fragments: [
        {
          kind: "paragraph",
          blockId: "p1",
          x: 72,
          y: 72,
          width: 672,
          height: 24,
          pmStart: 1,
          pmEnd: 20,
        },
        {
          kind: "paragraph",
          blockId: "p2",
          x: 72,
          y: 96,
          width: 672,
          height: 24,
          pmStart: 22,
          pmEnd: 42,
        },
      ],
    },
    {
      number: 2,
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
      size: { w: 816, h: 1056 },
      fragments: [
        {
          kind: "paragraph",
          blockId: "p3",
          x: 72,
          y: 72,
          width: 672,
          height: 24,
          pmStart: 44,
          pmEnd: 60,
        },
      ],
    },
    {
      number: 3,
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
      size: { w: 816, h: 1056 },
      // page-shell placeholder: no positioned content yet
      fragments: [],
    },
  ],
} satisfies Layout;

describe("getPageTextFromLayout", () => {
  test("joins a page's fragment text with newlines", () => {
    expect(getPageTextFromLayout(layout, doc, 1)).toBe("Page one first line\nPage one second line");
  });

  test("reads a later page independently", () => {
    expect(getPageTextFromLayout(layout, doc, 2)).toBe("Page two content");
  });

  test("returns an empty string for a page with no positioned fragments", () => {
    expect(getPageTextFromLayout(layout, doc, 3)).toBe("");
  });

  test("returns null for an out-of-range page", () => {
    expect(getPageTextFromLayout(layout, doc, 4)).toBeNull();
    expect(getPageTextFromLayout(layout, doc, 0)).toBeNull();
  });

  test("returns null for a non-integer page number", () => {
    expect(getPageTextFromLayout(layout, doc, 1.5)).toBeNull();
  });

  test("returns null when the layout hasn't been computed yet", () => {
    expect(getPageTextFromLayout(null, doc, 1)).toBeNull();
  });
});
