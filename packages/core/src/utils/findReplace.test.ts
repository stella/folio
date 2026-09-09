import { describe, expect, test } from "bun:test";

import type { Document } from "../types/document";
import { createDefaultFindOptions, findInDocument } from "./findReplace";

const document: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Outside stock" }] }],
        },
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [
                    {
                      type: "paragraph",
                      content: [{ type: "run", content: [{ type: "text", text: "Inside stock" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
};

describe("findInDocument", () => {
  test("finds matches in body and table paragraphs with stable paragraph offsets", () => {
    const matches = findInDocument(document, "stock", createDefaultFindOptions());

    expect(
      matches.map(({ paragraphIndex, startOffset, text }) => ({
        paragraphIndex,
        startOffset,
        text,
      })),
    ).toEqual([
      { paragraphIndex: 0, startOffset: 8, text: "stock" },
      { paragraphIndex: 1, startOffset: 7, text: "stock" },
    ]);
  });

  test("honors case and whole-word options", () => {
    const caseDocument: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Stock stock stockholder" }],
                },
              ],
            },
          ],
        },
      },
    };

    expect(
      findInDocument(caseDocument, "Stock", { matchCase: true, matchWholeWord: true }),
    ).toHaveLength(1);
    expect(
      findInDocument(caseDocument, "stock", { matchCase: false, matchWholeWord: true }),
    ).toHaveLength(2);
  });

  test("does not match text across an explicit page-break run", () => {
    const pageBreakDocument: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    { type: "text", text: "left" },
                    { type: "break", breakType: "page" },
                    { type: "text", text: "right" },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    expect(findInDocument(pageBreakDocument, "left", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(pageBreakDocument, "right", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(pageBreakDocument, "leftright", createDefaultFindOptions())).toEqual([]);
  });

  test.each([undefined, "textWrapping"] as const)(
    "counts a %s line-break type in search offsets",
    (breakType) => {
      const lineBreakDocument: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "text", text: "A" },
                      {
                        type: "break",
                        ...(breakType !== undefined ? { breakType } : {}),
                      },
                      { type: "text", text: "B" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };

      expect(findInDocument(lineBreakDocument, "A\nB", createDefaultFindOptions())).toHaveLength(1);
      expect(
        findInDocument(lineBreakDocument, "B", createDefaultFindOptions()).at(0),
      ).toMatchObject({ startOffset: 2 });
    },
  );
});
