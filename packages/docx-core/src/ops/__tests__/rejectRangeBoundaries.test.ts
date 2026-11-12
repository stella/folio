import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 } as const;

const documentWith = (content: ParagraphContent[]): Document => ({
  package: {
    document: { content: [{ type: "paragraph", paraId: "00000001", content }] },
  },
});

const insert = (
  document: Document,
  content: ParagraphContent[],
  id: number,
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at,
    slice: { content, openStart: 0, openEnd: 0 },
    revision: { id, author: "A", date: `2026-03-04T05:00:0${id}Z` },
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const insertAt = (
  document: Document,
  content: ParagraphContent[],
  offset: number,
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: { ...at, offset },
    slice: { content, openStart: 0, openEnd: 0 },
    // Ids for the pieces of a record the insertion cuts, such as a control.
    newIds: { revision: [900, 901], control: [910, 911] },
    revision: { id: 3, author: "A", date: "2026-03-04T05:00:03Z" },
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const rejectAll = (document: Document, ids: number[]): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const firstParagraph = (document: Document) => {
  const first = document.package.document.content[0];
  if (first?.type !== "paragraph") throw new Error("expected a paragraph");
  return first.content;
};

describe("rejecting a tracked insertion keeps the range boundaries it held", () => {
  test("the shrunk L3 counterexample", () => {
    const original = documentWith([{ type: "commentRangeStart", id: 37 }]);
    const once = insert(
      original,
      [
        {
          type: "hyperlink",
          rId: "rId9",
          href: "https://example.org",
          children: [],
        },
      ],
      1,
    );
    const twice = insert(
      once,
      [
        {
          type: "hyperlink",
          rId: "rId9",
          href: "https://example.org",
          children: [
            {
              type: "inlineWrapper",
              kind: "bidi",
              control: "embedding",
              content: [{ type: "commentRangeEnd", id: 37 }],
            },
          ],
        },
      ],
      2,
    );
    // The insertion put the comment's end before its start; rejecting it
    // keeps the end where it was, and the empty range closes on its start.
    expect(firstParagraph(rejectAll(twice, [1, 2]))).toEqual([
      { type: "commentRangeStart", id: 37 },
      { type: "commentRangeEnd", id: 37 },
    ]);
  });

  test("a kept start before a standing end moves to the end, never the end out of its revision", () => {
    // The end stands in an insertion this rejection leaves; the start the
    // rejection kept comes after it, so the start closes on the end.
    const document = documentWith([
      {
        type: "insertion",
        info: { id: 1, author: "A" },
        content: [{ type: "bookmarkEnd", id: 3 }],
      },
      {
        type: "insertion",
        info: { id: 2, author: "A" },
        content: [
          { type: "run", content: [{ type: "text", text: "x" }] },
          { type: "bookmarkStart", id: 3, name: "_Ref3" },
        ],
      },
    ]);

    expect(firstParagraph(rejectAll(document, [2]))).toEqual([
      {
        type: "insertion",
        info: { id: 1, author: "A" },
        content: [
          { type: "bookmarkStart", id: 3, name: "_Ref3" },
          { type: "bookmarkEnd", id: 3 },
        ],
      },
    ]);
  });

  test("a kept boundary that is no copy stays, though the paragraph holds it too", () => {
    const control = {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 41, tag: "clause" },
    } as const;
    const document = documentWith([
      { type: "bookmarkStart", id: 5, name: "_Ref5" },
      { type: "run", content: [{ type: "text", text: "a" }] },
      { type: "bookmarkEnd", id: 5 },
      {
        ...control,
        content: [
          {
            type: "insertion",
            info: { id: 2, author: "A" },
            content: [{ type: "bookmarkStart", id: 5, name: "_Ref5" }],
          },
        ],
      },
    ]);

    // The insertion's bookmark carries no copy ancestry: it is a boundary of
    // its own, so rejecting keeps it where the insertion was.
    expect(firstParagraph(rejectAll(document, [2]))).toEqual([
      { type: "bookmarkStart", id: 5, name: "_Ref5" },
      { type: "run", content: [{ type: "text", text: "a" }] },
      { type: "bookmarkEnd", id: 5 },
      {
        ...control,
        content: [{ type: "bookmarkStart", id: 5, name: "_Ref5" }],
      },
    ]);
  });

  // Accepting the copy makes it a boundary of its own, as inserting it
  // directly would, and rejecting the original's insertion keeps the
  // original, which is no copy: either order leaves both.
  test("rejecting a bookmark's insertion and accepting its copy's leaves the same, in either order", () => {
    const bookmarked = (text: string): ParagraphContent[] => [
      { type: "bookmarkStart", id: 7, name: "_Ref7" },
      { type: "run", content: [{ type: "text", text }] },
      { type: "bookmarkEnd", id: 7 },
    ];
    const base = documentWith([
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ]);
    const original = insertAt(base, bookmarked("x"), 1);
    const applied = applyDocumentOp(original, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: { ...at, offset: 3 },
      slice: { content: bookmarked("y"), openStart: 0, openEnd: 0 },
      revision: { id: 4, author: "A", date: "2026-03-04T05:00:04Z" },
    });
    if (applied.isErr()) throw applied.error;
    const tracked = applied.value.document;
    const accepted = (document: Document): Document => {
      const resolved = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [4],
        decision: REVISION_DECISIONS.ACCEPT,
      });
      if (resolved.isErr()) throw resolved.error;
      return resolved.value.document;
    };
    expect(firstParagraph(rejectAll(accepted(tracked), [3]))).toEqual(
      firstParagraph(accepted(rejectAll(tracked, [3]))),
    );
  });

  test("accepting a copy of a bookmark the paragraph holds leaves what inserting it directly does", () => {
    const bookmark = [
      { type: "bookmarkStart", id: 7, name: "_Ref7" },
      { type: "run", content: [{ type: "text", text: "x" }] },
      { type: "bookmarkEnd", id: 7 },
    ] as const satisfies ParagraphContent[];
    const base = documentWith([...bookmark]);
    const tracked = insertAt(base, structuredClone([...bookmark]), 1);
    const resolved = applyDocumentOp(tracked, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [3],
      decision: REVISION_DECISIONS.ACCEPT,
    });
    if (resolved.isErr()) throw resolved.error;
    const direct = applyDocumentOp(base, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: { ...at, offset: 1 },
      slice: {
        content: structuredClone([...bookmark]),
        openStart: 0,
        openEnd: 0,
      },
    });
    if (direct.isErr()) throw direct.error;
    const markers = (document: Document) =>
      firstParagraph(document).filter(
        ({ type }) => type === "bookmarkStart" || type === "bookmarkEnd",
      );
    expect(markers(resolved.value.document)).toEqual(
      markers(direct.value.document),
    );
    expect(markers(resolved.value.document)).toHaveLength(4);
  });

  test("a bookmark inside a rejected insertion stays where the insertion was", () => {
    const base = documentWith([
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ]);
    const tracked = insertAt(
      base,
      [{ type: "bookmarkStart", id: 5, name: "_Ref5" }],
      1,
    );
    expect(firstParagraph(rejectAll(tracked, [3]))).toEqual([
      { type: "run", content: [{ type: "text", text: "a" }] },
      { type: "bookmarkStart", id: 5, name: "_Ref5" },
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
  });

  test("a rejected copy of a boundary the paragraph still holds goes, so no id repeats", () => {
    const link = {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [
        { type: "run", content: [{ type: "text", text: "a" }] },
        { type: "bookmarkStart", id: 1, name: "_Link" },
      ],
    } as const satisfies ParagraphContent;
    const base = documentWith([link]);
    const tracked = insertAt(base, [structuredClone(link)], 0);
    expect(firstParagraph(rejectAll(tracked, [3]))).toEqual([link]);
  });

  test("a kept boundary goes inside the content control its insertion split, which joins again", () => {
    const base = documentWith([
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 2, tag: "clause" },
        content: [{ type: "run", content: [{ type: "text", text: "ab" }] }],
      },
    ]);
    const tracked = insertAt(
      base,
      [{ type: "moveToRangeStart", id: 1, name: "move1", author: "A" }],
      1,
    );
    const rejected = firstParagraph(rejectAll(tracked, [3]));
    expect(rejected.map(({ type }) => type)).toEqual(["inlineSdt"]);
    const [control] = rejected;
    if (control?.type !== "inlineSdt") throw new Error("expected the control");
    expect(control.properties.id).toBe(2);
    expect(JSON.stringify(control.content)).toContain('"moveToRangeStart"');
    expect(
      control.content
        .flatMap((item) => (item.type === "run" ? item.content : []))
        .map((content) => (content.type === "text" ? content.text : ""))
        .join(""),
    ).toBe("ab");
  });

  test("accepting a deletion removes the bookmark inside it, as deleting directly does", () => {
    const base = documentWith([
      {
        type: "deletion",
        info: { id: 4, author: "A", date: "2026-03-04T05:00:04Z" },
        content: [
          { type: "run", content: [{ type: "deletedText", text: "x" }] },
          { type: "bookmarkStart", id: 6, name: "_Ref6" },
        ],
      },
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
    const applied = applyDocumentOp(base, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [4],
      decision: REVISION_DECISIONS.ACCEPT,
    });
    if (applied.isErr()) throw applied.error;
    expect(firstParagraph(applied.value.document)).toEqual([
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
  });

  test("a boundary nested in rejected content is kept, the content around it goes", () => {
    const base = documentWith([
      { type: "commentRangeStart", id: 9 },
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ]);
    const tracked = insertAt(
      base,
      [
        {
          type: "hyperlink",
          rId: "rId9",
          href: "https://example.org",
          children: [{ type: "commentRangeEnd", id: 9 }],
        },
      ],
      1,
    );
    expect(firstParagraph(rejectAll(tracked, [3]))).toEqual([
      { type: "commentRangeStart", id: 9 },
      { type: "run", content: [{ type: "text", text: "a" }] },
      { type: "commentRangeEnd", id: 9 },
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
  });
});

describe("a join places a kept boundary only inside the pieces it merges", () => {
  // A paragraph split inside an insertion, its pieces either side of the
  // inserted mark, and a content control after the right piece.
  const tracked = (): Document => ({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "insertion",
                info: {
                  id: 303650,
                  author: "C",
                  date: "2026-03-04T05:00:02Z",
                },
                content: [
                  {
                    type: "run",
                    content: [
                      {
                        type: "text",
                        text: "x",
                      },
                    ],
                  },
                ],
                resolutionJoins: {
                  before: 0,
                  after: 3,
                  remove: 0,
                },
              },
              {
                type: "insertion",
                info: {
                  id: 303649,
                  author: "C",
                  date: "2026-03-04T05:00:01Z",
                },
                content: [
                  {
                    type: "run",
                    content: [
                      {
                        type: "text",
                        text: "x",
                      },
                    ],
                  },
                ],
                resolutionJoins: {
                  before: 0,
                  after: 0,
                  remove: 0,
                },
              },
            ],
            paraId: "00000001",
          },
          {
            type: "paragraph",
            paraId: "10000000",
            content: [
              {
                type: "run",
                content: [
                  {
                    type: "text",
                    text: "a",
                  },
                ],
              },
              {
                type: "insertion",
                info: {
                  id: 1,
                  author: "A",
                },
                content: [
                  {
                    type: "bookmarkEnd",
                    id: 3,
                  },
                ],
              },
            ],
            pPrMark: {
              kind: "ins",
              info: {
                id: 2,
                author: "C",
                date: "2026-03-04T05:00:00Z",
              },
              resolutionJoin: 1,
            },
          },
          {
            type: "paragraph",
            content: [
              {
                type: "insertion",
                info: {
                  id: 303648,
                  author: "A",
                },
                content: [
                  {
                    type: "run",
                    content: [
                      {
                        type: "text",
                        text: "a",
                      },
                    ],
                  },
                ],
              },
              {
                type: "inlineSdt",
                properties: {
                  sdtType: "richText",
                  id: 1,
                  tag: "clause",
                },
                content: [],
              },
            ],
            sectionProperties: {
              pageWidth: 11906,
            },
            paraId: "00000002",
          },
        ] satisfies BlockContent[],
      },
    },
  });
  const resolveIds = (document: Document, ids: number[]): Document => {
    const applied = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: ids,
      decision: REVISION_DECISIONS.REJECT,
    });
    if (applied.isErr()) throw applied.error;
    return applied.value.document;
  };
  const joined = (document: Document) =>
    document.package.document.content.at(-1);

  test("rejecting both pieces with the mark leaves the boundary standing beside the control", () => {
    const ids = [1, 2, 303648, 303649, 303650];
    const atOnce = joined(resolveIds(tracked(), ids));
    expect(atOnce).toMatchObject({
      type: "paragraph",
      content: [
        { type: "run", content: [{ type: "text", text: "a" }] },
        { type: "bookmarkEnd", id: 3 },
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 1, tag: "clause" },
          content: [],
        },
      ],
    });
    // The join happening first, or last, places it the same.
    for (const order of [
      [[2], [1], [303649, 303650]],
      [[1, 303648], [303649, 303650], [2]],
    ]) {
      let document = tracked();
      for (const step of order) document = resolveIds(document, step);
      expect(joined(document)).toEqual(atOnce);
    }
  });
});
