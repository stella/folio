import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
} from "../../packages/docx-core/src/model/document";
import { rejectedBoundaryHolderIdentities } from "./rejectedJoinSurvivors";

const bookmark = { type: "bookmarkStart", id: 1, name: "_Link" } as const;
const paragraph = (
  paraId: string,
  content: Paragraph["content"] = [],
): Paragraph => ({ type: "paragraph", paraId, content });
const added = (paraId: string, content: Paragraph["content"]): Paragraph => ({
  ...paragraph(paraId, content),
  pPrMark: { kind: "ins", info: { id: 1, author: "A" } },
});
const control = (content: BlockContent[]): BlockContent => ({
  type: "blockSdt",
  properties: { sdtType: "group", id: 9 },
  content,
});
const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const source = documentOf([
  added("00000001", [bookmark]),
  control([paragraph("00000002")]),
  paragraph("00000003"),
]);
const legacy = documentOf([
  paragraph("00000001", [bookmark]),
  control([paragraph("00000002")]),
  paragraph("00000003"),
]);

test("a paragraph kept for its bookmark is proven gone where the bookmark reaches the next paragraph's start", () => {
  expect(
    rejectedBoundaryHolderIdentities({
      source,
      actual: documentOf([
        control([paragraph("00000002")]),
        paragraph("00000003", [bookmark]),
      ]),
      expected: legacy,
    }),
  ).toStrictEqual(["00000001"]);
});

test("a bookmark that lands in any other paragraph proves nothing", () => {
  expect(
    rejectedBoundaryHolderIdentities({
      source,
      actual: documentOf([
        control([paragraph("00000002", [bookmark])]),
        paragraph("00000003"),
      ]),
      expected: legacy,
    }),
  ).toStrictEqual([]);
});
