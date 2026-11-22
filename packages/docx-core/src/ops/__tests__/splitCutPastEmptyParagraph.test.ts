/**
 * A split's cut goes on past an empty paragraph split off in it later: it
 * holds neither piece, so whichever of the two breaks is rejected first, the
 * pieces join.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const SPLIT = 1;
const EMPTY = 2;

const text = (value: string): Paragraph["content"][number] => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const control = (id: number, value: string): Paragraph["content"][number] => ({
  type: "inlineSdt",
  properties: { sdtType: "richText", id, tag: "clause" },
  content: [text(value)],
});

const bookmark = (id: number): BlockContent => ({
  type: "bookmarkStart",
  id,
  name: `_Ref${id}`,
});

const paragraph = (
  paraId: string,
  content: Paragraph["content"],
  mark?: Paragraph["pPrMark"],
): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
  ...(mark === undefined ? {} : { pPrMark: mark }),
});

const added = (
  id: number,
  date: string,
  resolutionJoin: number,
): NonNullable<Paragraph["pPrMark"]> => ({
  kind: "ins",
  info: { id, author: "A", date: `2026-03-04T05:00:0${date}Z` },
  resolutionJoin,
});

const rejected = (document: Document, ids: number[]): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const contentsOf = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph" ? block.content : block,
  );

const first = paragraph("00000001", [control(1, "a")], added(SPLIT, "1", 3));
const emptyAfter = paragraph("00000002", [], added(EMPTY, "2", 0));
const splitOff = paragraph("00000001", [control(1, "a")], added(EMPTY, "2", 0));
const emptyBefore = paragraph("00000002", [], added(SPLIT, "1", 3));
const last = paragraph("00000003", [control(2, "b")]);
const table: BlockContent = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        { type: "tableCell", content: [paragraph("00000009", [text("cell")])] },
      ],
    },
  ],
};
const blockControl = (content: Paragraph[]): BlockContent => ({
  type: "blockSdt",
  properties: { sdtType: "richText", id: 90 },
  content,
});

const rows: Record<string, BlockContent[]> = {
  "split off after the first piece": [first, emptyAfter, last],
  "split off before the second piece": [splitOff, emptyBefore, last],
  "with a range boundary before it": [first, bookmark(7), emptyAfter, last],
  "with a range boundary after it": [first, emptyAfter, bookmark(7), last],
  "with a table after the second piece": [
    first,
    emptyAfter,
    last,
    table,
    paragraph("0000000A", []),
  ],
  "with a section ending at the second piece": [
    first,
    emptyAfter,
    { ...last, sectionProperties: { pageWidth: 12_240 } },
    paragraph("0000000A", []),
  ],
  "inside a block control": [
    blockControl([first, emptyAfter, last]),
    paragraph("0000000A", []),
  ],
};

for (const [name, content] of Object.entries(rows))
  test(`a split's cut goes on past an empty paragraph, ${name}`, () => {
    const document: Document = { package: { document: { content } } };
    const atOnce = contentsOf(rejected(document, [SPLIT, EMPTY]));
    expect(contentsOf(rejected(rejected(document, [SPLIT]), [EMPTY]))).toEqual(
      atOnce,
    );
    expect(contentsOf(rejected(rejected(document, [EMPTY]), [SPLIT]))).toEqual(
      atOnce,
    );
  });

test("the pieces of the split are one again", () => {
  const document: Document = {
    package: { document: { content: [first, emptyAfter, last] } },
  };
  expect(contentsOf(rejected(rejected(document, [SPLIT]), [EMPTY]))).toEqual([
    [
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 1, tag: "clause" },
        content: [text("ab")],
      },
    ],
  ]);
});

test("an empty paragraph that was there before the split takes no cut", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("00000001", [control(1, "a")], added(SPLIT, "2", 3)),
          paragraph("00000002", [], added(EMPTY, "1", 0)),
          last,
        ],
      },
    },
  };
  expect(contentsOf(rejected(rejected(document, [SPLIT]), [EMPTY]))).toEqual([
    [control(1, "a"), control(2, "b")],
  ]);
});
