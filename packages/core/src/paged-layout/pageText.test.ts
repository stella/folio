import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { assertProperty } from "../../../../test/property-testing";
import { DISPLAY_MODES } from "../managers/EditorModeManager";
import type { DisplayMode } from "../managers/EditorModeManager";
import { projectMarkupView } from "../prosemirror/markupViewProjection";
import { schema as reviewSchema } from "../prosemirror/schema";
import { Schema } from "prosemirror-model";

import type { Layout } from "../layout-engine";
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
    expect(getPageTextFromLayout({ layout: layout, doc, page: 1 })).toBe(
      "Page one first line\nPage one second line",
    );
  });

  test("reads a later page independently", () => {
    expect(getPageTextFromLayout({ layout: layout, doc, page: 2 })).toBe("Page two content");
  });

  test("returns an empty string for a page with no positioned fragments", () => {
    expect(getPageTextFromLayout({ layout: layout, doc, page: 3 })).toBe("");
  });

  test("returns null for an out-of-range page", () => {
    expect(getPageTextFromLayout({ layout: layout, doc, page: 4 })).toBeNull();
    expect(getPageTextFromLayout({ layout: layout, doc, page: 0 })).toBeNull();
  });

  test("returns null for a non-integer page number", () => {
    expect(getPageTextFromLayout({ layout: layout, doc, page: 1.5 })).toBeNull();
  });

  test("returns null when the layout hasn't been computed yet", () => {
    expect(getPageTextFromLayout({ layout: null, doc, page: 1 })).toBeNull();
  });
});

test("page text follows each review projection for every inline revision span", async () => {
  await assertProperty(
    fc.property(
      fc.tuple(
        fc.string({ maxLength: 40 }),
        fc.string({ minLength: 1, maxLength: 40 }),
        fc.string({ minLength: 1, maxLength: 40 }),
        fc.string({ maxLength: 40 }),
      ),
      (texts) => {
        const kept = texts.at(0) ?? panic("Missing kept text");
        const inserted = texts.at(1) ?? panic("Missing inserted text");
        const deleted = texts.at(2) ?? panic("Missing deleted text");
        const tail = texts.at(3) ?? panic("Missing tail text");
        const insertion = (
          reviewSchema.marks["insertion"] ?? panic("Missing insertion mark")
        ).create({ revisionId: 1 });
        const deletion = (reviewSchema.marks["deletion"] ?? panic("Missing deletion mark")).create({
          revisionId: 2,
        });
        const runs = [
          reviewSchema.text(inserted, [insertion]),
          reviewSchema.text(deleted, [deletion]),
        ];
        if (kept) runs.unshift(reviewSchema.text(kept));
        if (tail) runs.push(reviewSchema.text(tail));
        const authored = reviewSchema.node("doc", null, [
          reviewSchema.node("paragraph", null, runs),
        ]);
        const state = EditorState.create({ doc: authored });
        const fixturePage = layout.pages.at(0) ?? panic("Missing fixture page");
        const projectedLayout = {
          pages: [
            {
              ...fixturePage,
              fragments: [
                {
                  kind: "paragraph",
                  blockId: "body",
                  x: 72,
                  y: 72,
                  width: 672,
                  height: 24,
                  pmStart: 1,
                  pmEnd: authored.content.size - 1,
                },
              ],
            },
          ],
        } satisfies Layout;
        const expected = {
          "all-markup": kept + inserted + deleted + tail,
          "simple-markup": kept + inserted + tail,
          "no-markup": kept + inserted + tail,
          original: kept + deleted + tail,
        } as const satisfies Record<DisplayMode, string>;
        for (const mode of DISPLAY_MODES) {
          expect(
            getPageTextFromLayout({
              layout: projectedLayout,
              doc: authored,
              page: 1,
              projection: projectMarkupView(state, mode),
            }),
          ).toBe(expected[mode]);
        }
        expect(state.doc).toBe(authored);
      },
    ),
    { numRuns: 40, id: "page text follows each review projection for every inline revision span" },
  );
});
