import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
} from "../types/document";
import { withoutOrphanCommentRanges } from "./commentRangeIntegrity";

const validComment = {
  id: 1,
  author: "Reviewer",
  content: [{ type: "paragraph", content: [] }],
} satisfies Comment;

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const paragraph = (content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  content,
});

const contentTypes = (block: BlockContent | undefined): string[] => {
  expect(block?.type).toBe("paragraph");
  if (block?.type !== "paragraph") {
    throw new Error("Expected paragraph");
  }
  return block.content.map((item) => item.type);
};

describe("withoutOrphanCommentRanges", () => {
  test("drops orphans across body, table, header, footnote, and endnote content", () => {
    const tableCellParagraph = paragraph([
      { type: "commentRangeEnd", id: 404 },
      run("table"),
      { type: "commentReference", id: 1 },
    ]);
    const document: Document = {
      package: {
        document: {
          comments: [validComment],
          content: [
            paragraph([
              { type: "commentRangeStart", id: 404 },
              run("body"),
              { type: "commentRangeEnd", id: 1 },
            ]),
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    {
                      type: "tableCell",
                      content: [tableCellParagraph],
                    },
                  ],
                },
              ],
            },
          ],
        },
        headers: new Map([
          [
            "rIdHeader",
            {
              type: "header",
              hdrFtrType: "default",
              content: [paragraph([{ type: "commentRangeStart", id: 405 }, run("header")])],
            },
          ],
        ]),
        footnotes: [
          {
            type: "footnote",
            id: 2,
            content: [paragraph([run("footnote"), { type: "commentRangeEnd", id: 406 }])],
          },
        ],
        endnotes: [
          {
            type: "endnote",
            id: 3,
            content: [paragraph([{ type: "commentReference", id: 407 }, run("endnote")])],
          },
        ],
      },
    };

    const result = withoutOrphanCommentRanges(document);

    expect(result).not.toBe(document);
    expect(contentTypes(result.package.document.content.at(0))).toEqual(["run", "commentRangeEnd"]);

    const table = result.package.document.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("Expected table");
    }
    expect(contentTypes(table.rows.at(0)?.cells.at(0)?.content.at(0))).toEqual([
      "run",
      "commentReference",
    ]);

    expect(contentTypes(result.package.headers?.get("rIdHeader")?.content.at(0))).toEqual(["run"]);
    expect(contentTypes(result.package.footnotes?.at(0)?.content.at(0))).toEqual(["run"]);
    expect(contentTypes(result.package.endnotes?.at(0)?.content.at(0))).toEqual(["run"]);
  });

  test("does not mutate the caller's document model", () => {
    const originalParagraph = paragraph([
      { type: "commentRangeStart", id: 404 },
      run("body"),
      { type: "commentReference", id: 404 },
    ]);
    const document: Document = {
      package: {
        document: {
          comments: [validComment],
          content: [originalParagraph],
        },
      },
    };

    const result = withoutOrphanCommentRanges(document);

    expect(contentTypes(document.package.document.content.at(0))).toEqual([
      "commentRangeStart",
      "run",
      "commentReference",
    ]);
    expect(contentTypes(result.package.document.content.at(0))).toEqual(["run"]);
    expect(result.package.document.content.at(0)).not.toBe(originalParagraph);
  });

  test("returns the same document reference when no orphan markers exist", () => {
    const document: Document = {
      package: {
        document: {
          comments: [validComment],
          content: [
            paragraph([
              { type: "commentRangeStart", id: 1 },
              run("body"),
              { type: "commentRangeEnd", id: 1 },
              { type: "commentReference", id: 1 },
            ]),
          ],
        },
      },
    };

    expect(withoutOrphanCommentRanges(document)).toBe(document);
  });

  test("drops orphan commentReference markers", () => {
    const document: Document = {
      package: {
        document: {
          comments: [validComment],
          content: [
            paragraph([run("before"), { type: "commentReference", id: 404 }, run("after")]),
          ],
        },
      },
    };

    const result = withoutOrphanCommentRanges(document);

    expect(contentTypes(result.package.document.content.at(0))).toEqual(["run", "run"]);
  });
});
