/**
 * A kept marker at a paragraph's edge waits on the pending paragraph mark
 * that can join it with its neighbour: across range boundaries standing
 * between the two paragraphs, and from inside a record at that edge.
 * Resolving something else leaves its cut standing, as joining the
 * paragraphs first does.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  KeptCut,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { childNodes, type InlineNode } from "../leaves";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const CUT = { keptFrom: "2026-03-04T05:00:01Z" } as const;
const MARK = {
  kind: "ins",
  info: { id: 9, author: "A", date: "2026-03-04T05:00:00Z" },
} as const;

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const inserted = (value: string): ParagraphContent => ({
  type: "insertion",
  info: { id: 7, author: "B", date: "2026-03-04T05:00:02Z" },
  content: [text(value)],
});

const keptEnd: ParagraphContent = { type: "bookmarkEnd", id: 1, keptCut: CUT };

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

/** The kept cuts of the bookmark ends in `nodes`, in document order. */
const cutsIn = (nodes: readonly InlineNode[]): (KeptCut | undefined)[] =>
  nodes.flatMap((node) =>
    node.type === "bookmarkEnd"
      ? [node.keptCut]
      : cutsIn(childNodes(node) ?? []),
  );

/** The kept cuts of the bookmark ends in `blocks`' paragraphs. */
const cutsInBlocks = (blocks: readonly BlockContent[]) =>
  blocks.flatMap((block) =>
    block.type === "paragraph" ? cutsIn(block.content) : [],
  );

const acceptedInsertion = (document: Document) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [7],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document.package.document.content;
};

test("a kept marker after a block-level boundary waits on the pending mark before it", () => {
  const content = acceptedInsertion(
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [{ type: "bookmarkStart", id: 1, name: "b" }, text("First")],
        pPrMark: MARK,
      },
      { type: "bookmarkStart", id: 5, name: "Clause" },
      {
        type: "paragraph",
        paraId: "00000002",
        content: [keptEnd, inserted("Second")],
      },
      { type: "bookmarkEnd", id: 5 },
    ]),
  );
  expect(cutsInBlocks(content)).toEqual([CUT]);
});

test("a kept marker ending a link waits on its paragraph's pending mark", () => {
  const content = acceptedInsertion(
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          { type: "bookmarkStart", id: 1, name: "b" },
          {
            type: "hyperlink",
            anchor: "target",
            children: [inserted("a"), keptEnd],
          },
        ],
        pPrMark: MARK,
      },
      { type: "paragraph", paraId: "00000002", content: [text("Second")] },
    ]),
  );
  expect(cutsInBlocks(content)).toEqual([CUT]);
});
