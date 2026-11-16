import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Paragraph, TextFormatting } from "../../types/document";
import { toProseDoc } from "./toProseDoc";

const emptyCarriers = {
  run: [{ type: "run", content: [] }],
  hyperlink: [{ type: "hyperlink", href: "https://example.org", children: [] }],
  insertion: [
    { type: "insertion", info: { id: 1, author: "Reviewer" }, content: [] },
  ],
  deletion: [
    { type: "deletion", info: { id: 1, author: "Reviewer" }, content: [] },
  ],
  nested: [
    {
      type: "insertion",
      info: { id: 1, author: "Reviewer" },
      content: [
        { type: "hyperlink", href: "https://example.org", children: [] },
      ],
    },
  ],
} satisfies Record<string, Paragraph["content"]>;

const settings = fc.record({
  paragraphStyleId: fc.option(fc.constantFrom("Normal", "Heading1"), {
    nil: undefined,
  }),
  mark: fc.record({
    styleId: fc.option(fc.constant("Emphasis"), { nil: undefined }),
    bold: fc.boolean(),
    fontSize: fc.integer({ min: 12, max: 48 }),
    fontFamily: fc.record({ ascii: fc.constantFrom("Arial", "Courier New") }),
  }),
  location: fc.constantFrom(
    "body" as const,
    "plainCell" as const,
    "styledCell" as const,
  ),
});

type SourceOptions = {
  content: Paragraph["content"];
  paragraphStyleId: string | undefined;
  mark: TextFormatting;
  location: "body" | "plainCell" | "styledCell";
};

const source = ({
  content,
  paragraphStyleId,
  mark,
  location,
}: SourceOptions) => {
  const paragraph = {
    type: "paragraph",
    paraId: "12345678",
    formatting: { styleId: paragraphStyleId, runProperties: mark },
    content,
  } satisfies Paragraph;
  const document = {
    package: {
      styles: {
        docDefaults: {
          rPr: { fontFamily: { ascii: "Calibri" }, fontSize: 22 },
        },
        styles: [
          {
            styleId: "Normal",
            type: "paragraph",
            default: true,
            rPr: { color: { rgb: "000000" } },
          },
          {
            styleId: "Heading1",
            type: "paragraph",
            basedOn: "Normal",
            rPr: { bold: true, fontSize: 28 },
          },
          { styleId: "Emphasis", type: "character", rPr: { italic: true } },
          {
            styleId: "CaretTable",
            type: "table",
            rPr: { color: { rgb: "112233" }, fontSize: 20 },
          },
        ],
      },
      document: {
        content:
          location === "body"
            ? [paragraph]
            : [
                {
                  type: "table",
                  ...(location === "styledCell"
                    ? { formatting: { styleId: "CaretTable" } }
                    : {}),
                  rows: [
                    {
                      type: "tableRow",
                      cells: [{ type: "tableCell", content: [paragraph] }],
                    },
                  ],
                },
              ],
      },
    },
  } satisfies Document;
  return document;
};

const caretFormatting = (document: Document) => {
  const prose = toProseDoc(document, { styles: document.package.styles });
  const first = prose.firstChild;
  const paragraph =
    first?.type.name === "table"
      ? first.firstChild?.firstChild?.firstChild
      : first;
  if (paragraph?.type.name !== "paragraph")
    panic("The fixture has one paragraph.");
  // Every declared carrier must actually emit an empty paragraph.
  expect(paragraph.childCount).toBe(0);
  return paragraph.attrs["defaultTextFormatting"];
};

test(
  "empty inline carriers have the caret formatting of a genuinely empty paragraph",
  () => {
    assertProperty(
      fc.property(settings, (picked) => {
        const expected = caretFormatting(source({ ...picked, content: [] }));
        for (const content of Object.values(emptyCarriers))
          expect(caretFormatting(source({ ...picked, content }))).toEqual(
            expected,
          );
      }),
      {
        id: "empty carrier caret formatting equals empty paragraph formatting",
        numRuns: 60,
        examples: [
          [
            {
              paragraphStyleId: "Heading1",
              mark: {
                styleId: "Emphasis",
                bold: false,
                fontSize: 22,
                fontFamily: { ascii: "Arial" },
              },
              location: "plainCell",
            },
          ],
        ],
      },
    );
  },
  propertyTestTimeout(30_000),
);
