import { describe, expect, test } from "bun:test";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Comment, DocumentBody, HeaderFooter } from "../types/document";
import { visitCommentMarkers } from "./commentAnchorIndex";
import { normalizeCommentReferences } from "./commentReferenceNormalization";

describe("comment reference normalization", () => {
  for (const level of ["table", "row"] as const) {
    test(`keeps a ${level} comment end after its last remaining child`, () => {
      const preserved = {
        children: [{ index: 2, xml: '<w:commentRangeEnd w:id="5"/>' }],
      };
      const documentBody: DocumentBody = {
        content: [
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [
                      {
                        type: "paragraph",
                        content: [
                          { type: "commentRangeStart", id: 5 },
                          {
                            type: "run",
                            content: [{ type: "text", text: "Covered" }],
                          },
                          { type: "commentReference", id: 5 },
                        ],
                      },
                    ],
                  },
                ],
                ...(level === "row" && { preserved }),
              },
            ],
            ...(level === "table" && { preserved }),
          },
          {
            type: "paragraph",
            content: [
              { type: "run", content: [{ type: "text", text: "Outside" }] },
            ],
          },
        ],
      };
      const comments: Comment[] = [{ id: 5, author: "Reviewer", content: [] }];
      expect(
        normalizeCommentReferences({ documentBody, comments })
          .reanchoredUnbalancedRanges,
      ).toBe(0);
      const markers: string[] = [];
      visitCommentMarkers(documentBody.content, ({ item }) =>
        markers.push(item.type),
      );
      expect(markers).toEqual([
        "commentRangeStart",
        "commentReference",
        "commentRangeEnd",
      ]);

      const covered: string[] = [];
      toProseDoc({
        package: { document: { ...documentBody, comments } },
      }).descendants((node) => {
        if (
          node.isText &&
          node.marks.some(({ type }) => type.name === "comment")
        )
          covered.push(node.textContent);
      });
      expect(covered).toEqual(["Covered"]);
    });
  }

  test("removes markers whose comments.xml entries are missing", () => {
    const documentBody: DocumentBody = {
      content: [
        {
          type: "paragraph",
          content: [
            { type: "commentRangeStart", id: 0 },
            {
              type: "run",
              content: [{ type: "text", text: "Visible" }],
            },
            { type: "commentRangeEnd", id: 0 },
            { type: "commentReference", id: 1 },
          ],
        },
      ],
    };
    const comments: Comment[] = [
      {
        id: 1,
        author: "Reviewer",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "commentReference", id: 404 },
              {
                type: "run",
                content: [{ type: "text", text: "Comment body" }],
              },
            ],
          },
        ],
      },
    ];
    const headers = new Map<string, HeaderFooter>([
      [
        "rIdHeader",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              content: [{ type: "commentReference", id: 404 }],
            },
          ],
        },
      ],
    ]);

    const result = normalizeCommentReferences({
      documentBody,
      comments,
      headers,
    });

    expect(result).toEqual({
      removedDanglingReferences: 4,
      reanchoredUnbalancedRanges: 0,
      droppedUnpairedBoundaries: 0,
      droppedDuplicateReferences: 0,
    });
    const paragraph = documentBody.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      return;
    }
    expect(paragraph.content.map((content) => content.type)).toEqual([
      "run",
      "commentReference",
    ]);
    expect(
      comments[0]?.content[0]?.content.map((content) => content.type),
    ).toEqual(["run"]);
    expect(
      headers
        .get("rIdHeader")
        ?.content[0]?.content.map((content) => content.type),
    ).toEqual([]);
  });

  test("re-anchors unbalanced valid ranges as point comments", () => {
    const documentBody: DocumentBody = {
      comments: [
        {
          id: 0,
          author: "Reviewer",
          content: [{ type: "paragraph", content: [] }],
        },
      ],
      content: [
        {
          type: "paragraph",
          content: [
            { type: "commentRangeStart", id: 0 },
            {
              type: "run",
              content: [{ type: "text", text: "Text" }],
            },
          ],
        },
      ],
    };

    const result = normalizeCommentReferences({
      documentBody,
      comments: documentBody.comments ?? [],
    });

    expect(result).toEqual({
      removedDanglingReferences: 0,
      reanchoredUnbalancedRanges: 1,
      droppedUnpairedBoundaries: 0,
      droppedDuplicateReferences: 0,
    });
    const paragraph = documentBody.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      return;
    }
    expect(paragraph.content.at(0)).toEqual({
      type: "commentReference",
      id: 0,
    });
  });

  test("re-anchors out-of-order valid range markers as one point comment", () => {
    const documentBody: DocumentBody = {
      comments: [
        {
          id: 0,
          author: "Reviewer",
          content: [{ type: "paragraph", content: [] }],
        },
      ],
      content: [
        {
          type: "paragraph",
          content: [
            { type: "commentRangeEnd", id: 0 },
            {
              type: "run",
              content: [{ type: "text", text: "Text" }],
            },
            { type: "commentRangeStart", id: 0 },
          ],
        },
      ],
    };

    const result = normalizeCommentReferences({
      documentBody,
      comments: documentBody.comments ?? [],
    });

    expect(result).toEqual({
      removedDanglingReferences: 0,
      reanchoredUnbalancedRanges: 1,
      droppedUnpairedBoundaries: 1,
      droppedDuplicateReferences: 0,
    });
    const paragraph = documentBody.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") {
      return;
    }
    expect(paragraph.content.map((content) => content.type)).toEqual([
      "commentReference",
      "run",
    ]);
  });

  test("drops a lone boundary of a comment its reference anchors", () => {
    const comments: Comment[] = [{ id: 7, author: "Reviewer", content: [] }];
    const documentBody: DocumentBody = {
      comments,
      content: [
        { type: "commentRangeStart", id: 7 },
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Text" }] },
            { type: "commentRangeEnd", id: 7 },
            { type: "commentRangeEnd", id: 7 },
            { type: "commentReference", id: 7 },
          ],
        },
        { type: "commentRangeEnd", id: 7 },
      ],
    };

    const result = normalizeCommentReferences({ documentBody, comments });

    expect(result).toEqual({
      removedDanglingReferences: 0,
      reanchoredUnbalancedRanges: 0,
      droppedUnpairedBoundaries: 2,
      droppedDuplicateReferences: 0,
    });
    expect(documentBody.content.at(-1)).toEqual({
      type: "preservedBlock",
      xml: '<w:commentRangeEnd w:id="7"/>',
    });
    const paragraph = documentBody.content.at(1);
    if (paragraph?.type !== "paragraph") throw new Error("expected paragraph");
    expect(paragraph.content.map((content) => content.type)).toEqual([
      "run",
      "commentRangeEnd",
      "commentReference",
    ]);
  });
});
