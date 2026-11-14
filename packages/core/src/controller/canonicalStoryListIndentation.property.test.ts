/**
 * A list paragraph projects the same list layout in every story.
 *
 * Headers, footers and notes are converted without the body's document, so a
 * story conversion reads the package's numbering only when it is given it.
 * Without it a numbered paragraph keeps its cached marker but loses the
 * indentation its level gives, and a canonical editor painted header lists
 * flush left. The same numbered paragraphs, placed in the body and in every
 * other story, must read the same list attributes in each story's canonical
 * projection as in the body's.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { OP_STORIES, type OpStory } from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document, Paragraph } from "../types/document";

type Item = {
  numbering: "direct" | "style";
  ilvl: number;
  /** A left indentation the paragraph states itself, over its level's. */
  indentLeft: number | null;
};

const itemArbitrary: fc.Arbitrary<Item> = fc.record({
  numbering: fc.constantFrom("direct" as const, "style" as const),
  ilvl: fc.nat({ max: 2 }),
  indentLeft: fc.option(fc.constantFrom(0, 360, 2160), { nil: null }),
});

/** Copies of the same paragraphs, so the stories share no record. */
const paragraphs = (items: readonly Item[]): Paragraph[] =>
  items.map(({ numbering, ilvl, indentLeft }, index) => ({
    type: "paragraph",
    formatting: {
      ...(numbering === "direct"
        ? { numPr: { kind: "reference", numId: 1, ilvl } }
        : { styleId: "ListStyle" }),
      ...(indentLeft === null ? {} : { indentLeft }),
    },
    content: [
      {
        type: "run",
        content: [{ type: "text", text: `Item ${String(index)}` }],
      },
    ],
  }));

const page = {
  pageWidth: 12240,
  pageHeight: 15840,
  headerReferences: [{ type: "default" as const, rId: "rIdHeader" }],
  footerReferences: [{ type: "default" as const, rId: "rIdFooter" }],
};

const listsInEveryStory = (items: readonly Item[]): Document => {
  const body = paragraphs(items);
  body.push({
    type: "paragraph",
    content: [
      {
        type: "run",
        content: [
          { type: "footnoteRef", id: 1 },
          { type: "endnoteRef", id: 1 },
        ],
      },
    ],
  });
  return {
    package: {
      document: { content: body, finalSectionProperties: page },
      headers: new Map([
        [
          "rIdHeader",
          { type: "header", hdrFtrType: "default", content: paragraphs(items) },
        ],
      ]),
      footers: new Map([
        [
          "rIdFooter",
          { type: "footer", hdrFtrType: "default", content: paragraphs(items) },
        ],
      ]),
      footnotes: [{ type: "footnote", id: 1, content: paragraphs(items) }],
      endnotes: [{ type: "endnote", id: 1, content: paragraphs(items) }],
      styles: {
        styles: [
          {
            styleId: "ListStyle",
            type: "paragraph",
            pPr: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
          },
        ],
      },
      numbering: {
        abstractNums: [
          {
            abstractNumId: 0,
            levels: [0, 1, 2].map((ilvl) => ({
              ilvl,
              start: 1,
              numFmt: "decimal" as const,
              lvlText: `%${String(ilvl + 1)}.`,
              pPr: {
                indentLeft: 720 * (ilvl + 1),
                indentFirstLine: -360,
                hangingIndent: true,
              },
            })),
          },
        ],
        nums: [{ numId: 1, abstractNumId: 0 }],
      },
    },
  };
};

const STORIES: readonly OpStory[] = [
  { kind: "header", rId: "rIdHeader" },
  { kind: "footer", rId: "rIdFooter" },
  { kind: "footnote", id: 1 },
  { kind: "endnote", id: 1 },
];

/** The list layout a paragraph projects, without its counter. */
const LIST_LAYOUT = [
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
  "numPr",
  "numPrFromStyle",
  "listNumFmt",
  "listIsBullet",
] as const;

const listLayouts = (doc: PMNode, count: number) => {
  const layouts: Record<string, unknown>[] = [];
  doc.forEach((node) => {
    if (layouts.length >= count || node.type.name !== "paragraph") return;
    layouts.push(
      Object.fromEntries(LIST_LAYOUT.map((key) => [key, node.attrs[key]])),
    );
  });
  return layouts;
};

const expectSameListLayoutInEveryStory = async (items: readonly Item[]) => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(listsInEveryStory(items)),
  );
  if (prepared.isErr()) throw prepared.error;
  const parsed = await parseDocx(prepared.value, {
    detectVariables: false,
    preloadFonts: false,
  });
  const created = createCanonicalSession(parsed, parsed.package.styles);
  if (created.isErr()) throw created.error;
  const session = created.value;
  const body = listLayouts(
    session.projectStory(OP_STORIES.MAIN).unwrap().doc,
    items.length,
  );
  expect(body).toHaveLength(items.length);
  for (const story of STORIES) {
    const projected = session.projectStory(story);
    if (projected.isErr()) throw projected.error;
    // A note's first paragraph also holds the note's mark; layout is the paragraph's.
    expect({
      story,
      layouts: listLayouts(projected.value.doc, items.length),
    }).toEqual({ story, layouts: body });
  }
};

/** Both numbering sources at every level, as the pinned header and footer cases. */
const PINNED_ITEMS: readonly Item[] = [
  { numbering: "direct", ilvl: 0, indentLeft: null },
  { numbering: "direct", ilvl: 1, indentLeft: null },
  { numbering: "style", ilvl: 0, indentLeft: null },
  { numbering: "direct", ilvl: 2, indentLeft: 360 },
];

test("pinned header, footer and note lists project the body's list layout", async () => {
  await expectSameListLayoutInEveryStory(PINNED_ITEMS);
});

test(
  "every story projects a list paragraph as the body does",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(itemArbitrary, { minLength: 1, maxLength: 6 }),
        expectSameListLayoutInEveryStory,
      ),
      {
        id: "every story projects a list paragraph as the body does",
        numRuns: 40,
      },
    );
  },
  propertyTestTimeout(30_000),
);
