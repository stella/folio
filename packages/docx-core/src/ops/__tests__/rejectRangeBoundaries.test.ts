import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
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
