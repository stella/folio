import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import * as blocks from "../blocks";
import {
  createEditorIntentIdAllocator,
  type EditorIntent,
} from "../editorIntent";
import { formatParaId } from "../idLease";
import { findStoryParagraph } from "../paragraphLocator";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const paragraph = (paraId: string | undefined, leaves = 1): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: [
    {
      type: "run",
      content: Array.from({ length: leaves }, () => ({
        type: "text",
        text: "x",
      })),
    },
  ],
});

const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const replaceAt = (blockId: string) =>
  ({
    type: "replaceText",
    from: { story: OP_STORIES.MAIN, blockId, offset: 0 },
    to: { story: OP_STORIES.MAIN, blockId: blockId.toUpperCase(), offset: 3 },
    text: "y",
  }) satisfies EditorIntent;

test("structural demand skips the generic selection and single endpoints use indexed lookup", () => {
  const content = Array.from({ length: 256 }, (_, index) =>
    paragraph(formatParaId(index + 1)),
  );
  const document = documentOf(content);
  const body = document.package.document;
  const anchor = "000000ff";
  const insert = {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: anchor },
    paragraphs: [paragraph(undefined)],
  } as const satisfies EditorIntent;
  const table = {
    type: "table",
    operation: {
      type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
      story: OP_STORIES.MAIN,
      blockId: anchor,
      patch: {},
    },
  } as const satisfies EditorIntent;
  const allocate = createEditorIntentIdAllocator();
  expect(findStoryParagraph(body, anchor)?.paragraph.paraId).toBe("000000FF");

  const walks = spyOn(blocks, "storyParagraphs");
  try {
    expect(allocate(document, insert).newIds.revision).toHaveLength(14);
    expect(allocate(document, replaceAt(anchor)).newIds.revision).toHaveLength(
      12,
    );
    expect(walks.mock.calls.some(([seen]) => seen === body)).toBe(false);

    expect(createEditorIntentIdAllocator()(document, table).newIds).toEqual({
      revision: [2],
      control: [],
    });
    // The table locator still reads its story once.
    expect(walks.mock.calls.filter(([seen]) => seen === body)).toHaveLength(1);
  } finally {
    walks.mockRestore();
  }

  expect(() =>
    allocate(document, {
      ...insert,
      story: { kind: "header", rId: "rIdMissing" },
    }),
  ).toThrow("An operation must address an existing document story.");
});

test("allocation keeps the first document-order match and the missing-endpoint budget", () => {
  const allocate = createEditorIntentIdAllocator();
  const document = documentOf([
    paragraph("000000aa", 2),
    paragraph(undefined, 1),
    paragraph("000000AA", 3),
    paragraph("", 3),
  ]);

  expect(
    allocate(document, replaceAt("000000aA")).newIds.revision,
  ).toHaveLength(20);
  expect(allocate(document, replaceAt("")).newIds.revision).toHaveLength(12);
  expect(allocate(document, replaceAt("000000AB")).newIds).toEqual({
    revision: [2],
    control: [1],
  });
});
