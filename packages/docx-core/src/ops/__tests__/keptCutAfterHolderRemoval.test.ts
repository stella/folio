/**
 * A kept marker's cut waits on a pending paragraph mark only while a paragraph
 * follows for it to join: where the resolution removes the paragraph that
 * stood only for the insertions it rejects, the cut is settled without it.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { settledParagraphKeptCuts } from "../keptCut";
import { storyBody } from "../stories";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const info = (id: number, second: number) => ({
  id,
  author: "A",
  date: `2026-03-04T05:00:0${second}Z`,
});

const insertion = (
  id: number,
  second: number,
  content: ParagraphContent[],
): ParagraphContent => ({
  type: "insertion",
  info: info(id, second),
  content,
});

const paragraph = (
  paraId: string,
  content: ParagraphContent[],
  mark?: number,
): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
  ...(mark === undefined
    ? {}
    : { pPrMark: { kind: "ins", info: info(mark, mark) } }),
});

const documentOf = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", [
          { type: "bookmarkStart", id: 1, name: "_Ref1" },
          text("o"),
        ]),
        {
          type: "blockSdt",
          properties: { sdtType: "richText" },
          content: [
            paragraph(
              "00000002",
              [
                insertion(6, 6, [text("z")]),
                text("a"),
                {
                  type: "bookmarkEnd",
                  id: 1,
                  keptCut: { keptFrom: "2026-03-04T05:00:01Z" },
                },
              ],
              7,
            ),
            paragraph("00000003", [insertion(5, 5, [text("x")])], 8),
          ],
        } satisfies BlockContent,
        paragraph("00000004", [text("p")]),
      ],
    },
  },
});

const rejected = (document: Document, revisionIds: number[]): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const unsettled = (document: Document) =>
  settledParagraphKeptCuts(storyBody(document, OP_STORIES.MAIN));

test("a kept cut is settled where removing an emptied paragraph leaves its own", () => {
  const held = rejected(documentOf(), [8]);
  const result = rejected(held, [5, 6]);
  const control = result.package.document.content.find(
    (block) => block.type === "blockSdt",
  );
  expect(control?.type === "blockSdt" && control.content.length).toBe(1);
  expect(unsettled(result)).toEqual([]);
});
