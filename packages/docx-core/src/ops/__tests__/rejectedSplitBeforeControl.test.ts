/**
 * A paragraph split while its own break and a property change are pending,
 * before a block content control that ends the body: rejecting everything at
 * once gives back what rejecting it before the split does. The paragraph the
 * rejected split joins into it is no paragraph that stays beside it, also
 * with a range boundary between the two and the paragraphs before.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const SPLIT = 7;

const empty = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [],
});

const added = (paraId: string, id: number): Paragraph => ({
  ...empty(paraId),
  pPrMark: { kind: "ins", info: { id, author: "A" } },
});

const changed = (paraId: string): Paragraph => ({
  ...added(paraId, 5),
  content: [{ type: "bookmarkStart", id: 1, name: "_Ref1" }],
  propertyChanges: [
    {
      type: "paragraphPropertyChange",
      info: { id: 4, author: "A" },
      previousFormatting: { alignment: "end" },
    },
  ],
});

const control: BlockContent = {
  type: "blockSdt",
  properties: { sdtType: "richText" },
  content: [empty("00000005"), empty("00000006")],
};

const splitInTwo: Paragraph = {
  ...empty("00000004"),
  pPrMark: {
    kind: "ins",
    info: { id: SPLIT, author: "C", date: "2026-03-04T05:00:01Z" },
    resolutionJoin: 0,
    survivor: "first",
  },
};

const rejected = (
  content: BlockContent[],
  revisionIds: number[],
): BlockContent[] => {
  const document: Document = { package: { document: { content } } };
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document.package.document.content;
};

const before: Record<string, BlockContent[]> = {
  "first in the body": [],
  "after a paragraph whose break is rejected too": [added("00000003", 3)],
  "after such a paragraph and a range boundary": [
    added("00000003", 3),
    { type: "bookmarkStart", id: 90, name: "_GoBack" },
  ],
};

for (const [name, blocks] of Object.entries(before))
  test(`a rejected split before a content control leaves the paragraph as it was (${name})`, () => {
    const ids = blocks.some((block) => block.type === "paragraph")
      ? [3, 4, 5]
      : [4, 5];
    const unsplit = rejected([...blocks, changed("00000004"), control], ids);
    expect(
      rejected(
        [...blocks, splitInTwo, changed("10000000"), control],
        [...ids, SPLIT],
      ),
    ).toEqual(unsplit);
    expect(unsplit.at(0)).toMatchObject({
      type: "paragraph",
      formatting: { alignment: "end" },
    });
  });
