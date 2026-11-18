/**
 * A reply's markers stand beside its thread's markers, inside the hyperlink,
 * content control or wrapper they stand in: a reply cuts none of them.
 */

import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES } from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const CONTAINERS: Record<
  string,
  (content: ParagraphContent[]) => ParagraphContent
> = {
  hyperlink: (children) => ({
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children,
  }),
  "content control": (content) => ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 5, tag: "clause" },
    content,
  }),
  "bidirectional wrapper": (content) => ({
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content,
  }),
};

const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "00000002", content }],
      comments: [
        {
          id: 7,
          author: "A",
          content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
        },
        {
          id: 8,
          author: "A",
          content: [{ type: "paragraph", paraId: "00000003", content: [] }],
        },
      ],
    },
  },
});

const reply = {
  type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
  comment: {
    id: 9,
    author: "B",
    done: false,
    date: "2026-01-02T03:04:05Z",
    content: [{ type: "paragraph", paraId: "7AAAAAAA", content: [] }],
  },
  anchor: { kind: "reply", parentId: 8 },
} as const;

describe.each(Object.entries(CONTAINERS))("in a %s", (_, container) => {
  test.each([
    {
      name: "a range over text",
      inner: [
        { type: "commentRangeStart", id: 8 },
        run("abc"),
        { type: "commentRangeEnd", id: 8 },
        { type: "commentReference", id: 8 },
      ],
      replied: [
        { type: "commentRangeStart", id: 8 },
        { type: "commentRangeStart", id: 9 },
        run("abc"),
        { type: "commentRangeEnd", id: 8 },
        { type: "commentRangeEnd", id: 9 },
        { type: "commentReference", id: 8 },
        { type: "commentReference", id: 9 },
      ],
    },
    {
      name: "a collapsed range beside another thread's reference",
      inner: [
        { type: "commentRangeStart", id: 8 },
        { type: "commentRangeEnd", id: 8 },
        { type: "commentReference", id: 7 },
      ],
      replied: [
        { type: "commentRangeStart", id: 8 },
        { type: "commentRangeStart", id: 9 },
        { type: "commentRangeEnd", id: 8 },
        { type: "commentRangeEnd", id: 9 },
        { type: "commentReference", id: 7 },
      ],
    },
  ] satisfies {
    name: string;
    inner: ParagraphContent[];
    replied: ParagraphContent[];
  }[])("a reply to $name keeps the container whole", ({ inner, replied }) => {
    const document = documentOf([run("x"), container(inner), run("y")]);
    const applied = applyDocumentOps(document, [reply]);
    if (applied.isErr()) throw applied.error;
    const [paragraph] = applied.value.document.package.document.content;
    expect(paragraph).toMatchObject({
      content: [run("x"), container(replied), run("y")],
    });
    const undone = applyDocumentOps(
      applied.value.document,
      applied.value.inverse,
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
  });
});
