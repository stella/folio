import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
} from "../types/document";
import {
  patchBreaksCommentRangeBalance,
  withoutOrphanCommentRanges,
} from "./commentRangeIntegrity";
import { visitDocxContentSlots } from "./paragraphTraversal";

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
              content: [
                paragraph([
                  { type: "commentRangeStart", id: 405 },
                  run("header"),
                ]),
              ],
            },
          ],
        ]),
        footnotes: [
          {
            type: "footnote",
            id: 2,
            content: [
              paragraph([
                run("footnote"),
                { type: "commentRangeEnd", id: 406 },
              ]),
            ],
          },
        ],
        endnotes: [
          {
            type: "endnote",
            id: 3,
            content: [
              paragraph([
                { type: "commentReference", id: 407 },
                run("endnote"),
              ]),
            ],
          },
        ],
      },
    };

    const result = withoutOrphanCommentRanges(document);

    expect(result).not.toBe(document);
    expect(contentTypes(result.package.document.content.at(0))).toEqual([
      "run",
      "commentRangeEnd",
    ]);

    const table = result.package.document.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("Expected table");
    }
    expect(contentTypes(table.rows.at(0)?.cells.at(0)?.content.at(0))).toEqual([
      "run",
      "commentReference",
    ]);

    expect(
      contentTypes(result.package.headers?.get("rIdHeader")?.content.at(0)),
    ).toEqual(["run"]);
    expect(
      contentTypes(result.package.footnotes?.at(0)?.content.at(0)),
    ).toEqual(["run"]);
    expect(contentTypes(result.package.endnotes?.at(0)?.content.at(0))).toEqual(
      ["run"],
    );
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
    expect(contentTypes(result.package.document.content.at(0))).toEqual([
      "run",
    ]);
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

  test("drops orphan markers inside comment bodies", () => {
    const document: Document = {
      package: {
        document: {
          comments: [
            {
              ...validComment,
              content: [
                paragraph([
                  { type: "commentRangeStart", id: 404 },
                  run("comment body"),
                  { type: "commentReference", id: 1 },
                ]),
              ],
            },
          ],
          content: [paragraph([run("body")])],
        },
      },
    };

    const result = withoutOrphanCommentRanges(document);

    expect(result).not.toBe(document);
    const commentParagraph = result.package.document.comments
      ?.at(0)
      ?.content.at(0);
    expect(commentParagraph?.content.map((item) => item.type)).toEqual([
      "run",
      "commentReference",
    ]);
  });

  test("drops orphan commentReference markers", () => {
    const document: Document = {
      package: {
        document: {
          comments: [validComment],
          content: [
            paragraph([
              run("before"),
              { type: "commentReference", id: 404 },
              run("after"),
            ]),
          ],
        },
      },
    };

    const result = withoutOrphanCommentRanges(document);

    expect(contentTypes(result.package.document.content.at(0))).toEqual([
      "run",
      "run",
    ]);
  });
});

type CommentMarker = Extract<
  ParagraphContent,
  {
    type: "commentRangeStart" | "commentRangeEnd" | "commentReference";
  }
>;

const MARKER_TYPES = [
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
] as const satisfies readonly CommentMarker["type"][];

const revision = { id: 50, author: "Reviewer" };

/** Every inline container a comment marker can sit in, holding one marker. */
const MARKER_CONTAINERS: Record<
  string,
  (marker: CommentMarker) => ParagraphContent
> = {
  "content control": (marker) => ({
    type: "inlineSdt",
    properties: { sdtType: "richText" },
    content: [marker],
  }),
  hyperlink: (marker) => ({
    type: "hyperlink",
    href: "https://example.com",
    children: [marker],
  }),
  insertion: (marker) => ({
    type: "insertion",
    info: revision,
    content: [marker],
  }),
  deletion: (marker) => ({
    type: "deletion",
    info: revision,
    content: [marker],
  }),
  "move from": (marker) => ({
    type: "moveFrom",
    info: revision,
    content: [marker],
  }),
  "move to": (marker) => ({
    type: "moveTo",
    info: revision,
    content: [marker],
  }),
  "bidi wrapper": (marker) => ({
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content: [marker],
  }),
  "smart tag": (marker) => ({
    type: "inlineWrapper",
    kind: "smartTag",
    element: "place",
    content: [marker],
  }),
  "custom XML wrapper": (marker) => ({
    type: "inlineWrapper",
    kind: "customXml",
    element: "clause",
    content: [marker],
  }),
  // A simple field admits no marker child of its own; one sits in a wrapper
  // the field holds.
  "simple field": (marker) => ({
    type: "simpleField",
    instruction: " PAGE ",
    fieldType: "PAGE",
    content: [
      {
        type: "inlineWrapper",
        kind: "customXml",
        element: "clause",
        content: [marker],
      },
    ],
  }),
  "hyperlink inside a content control": (marker) => ({
    type: "inlineSdt",
    properties: { sdtType: "richText" },
    content: [
      { type: "hyperlink", href: "https://example.com", children: [marker] },
    ],
  }),
  "text box of a shape in a run": (marker) => ({
    type: "run",
    content: [
      {
        type: "shape",
        shape: {
          type: "shape",
          shapeType: "rect",
          size: { width: 914400, height: 914400 },
          textBody: { content: [paragraph([run("boxed"), marker])] },
        },
      },
    ],
  }),
};

/** Every comment marker id the package's stories hold, at every depth. */
const commentMarkerIds = (document: Document): number[] => {
  const ids: number[] = [];
  visitDocxContentSlots(
    {
      documentBody: document.package.document,
      headers: document.package.headers,
      footers: document.package.footers,
      footnotes: document.package.footnotes,
      endnotes: document.package.endnotes,
    },
    (slot) => {
      if (
        slot.item.type === "commentRangeStart" ||
        slot.item.type === "commentRangeEnd" ||
        slot.item.type === "commentReference"
      ) {
        ids.push(slot.item.id);
      }
    },
  );
  return ids;
};

const documentWithContainer = (container: ParagraphContent): Document => ({
  package: {
    document: {
      comments: [validComment],
      content: [
        paragraph([run("before"), container, run("after")]),
        paragraph([run("untouched")]),
      ],
    },
  },
});

describe("withoutOrphanCommentRanges inside inline containers", () => {
  for (const [containerName, contain] of Object.entries(MARKER_CONTAINERS)) {
    for (const markerType of MARKER_TYPES) {
      test(`drops an orphan ${markerType} inside a ${containerName}`, () => {
        const document = documentWithContainer(
          contain({ type: markerType, id: 404 }),
        );
        expect(commentMarkerIds(document)).toEqual([404]);

        const result = withoutOrphanCommentRanges(document);

        expect(commentMarkerIds(result)).toEqual([]);
        // The container stays; only the marker goes.
        expect(contentTypes(result.package.document.content.at(0))).toEqual(
          contentTypes(document.package.document.content.at(0)),
        );
        // A paragraph holding no orphan is the same record.
        expect(result.package.document.content.at(1)).toBe(
          document.package.document.content.at(1),
        );
      });

      test(`keeps a held ${markerType} inside a ${containerName}`, () => {
        const document = documentWithContainer(
          contain({ type: markerType, id: validComment.id }),
        );

        expect(withoutOrphanCommentRanges(document)).toBe(document);
      });
    }
  }
});

const story = (...paragraphs: string[]): string =>
  paragraphs.map((content) => `<w:p w14:paraId="1">${content}</w:p>`).join("");

const OPENS = '<w:commentRangeStart w:id="7"/>';
const CLOSES = '<w:commentRangeEnd w:id="7"/>';

describe("patchBreaksCommentRangeBalance", () => {
  test("refuses a patch that drops a range's only start", () => {
    expect(
      patchBreaksCommentRangeBalance(
        story(`${OPENS}<w:r/>`, `<w:r/>${CLOSES}`),
        story("<w:r/>", `<w:r/>${CLOSES}`),
      ),
    ).toBe(true);
  });

  test("refuses a patch that leaves a start without its end", () => {
    expect(
      patchBreaksCommentRangeBalance(
        story(`${OPENS}<w:r/>`, `<w:r/>${CLOSES}`),
        story(`${OPENS}<w:r/>`, "<w:r/>"),
      ),
    ).toBe(true);
  });

  test("refuses a patch that moves a start after its end", () => {
    expect(
      patchBreaksCommentRangeBalance(
        story(`${OPENS}<w:r/>`, `<w:r/>${CLOSES}`),
        story(`<w:r/>${CLOSES}`, `${OPENS}<w:r/>`),
      ),
    ).toBe(true);
  });

  test("allows a balanced patch, and one that only carries the source's own imbalance", () => {
    expect(
      patchBreaksCommentRangeBalance(
        story(`${OPENS}<w:r/>`, `<w:r/>${CLOSES}`),
        story(`${OPENS}<w:r>x</w:r>`, `<w:r/>${CLOSES}`),
      ),
    ).toBe(false);
    expect(
      patchBreaksCommentRangeBalance(
        story(`<w:r/>${CLOSES}`),
        story(`<w:r>x</w:r>${CLOSES}`),
      ),
    ).toBe(false);
  });
});
