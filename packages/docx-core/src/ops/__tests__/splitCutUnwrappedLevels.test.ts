/**
 * A split's cut through a tracked change ends at that change's level once a
 * resolution unwraps its pieces: rejoining the paragraphs then merges no
 * records the split never cut.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { childNodes, type InlineNode } from "../leaves";
import { isTrackedWrapper } from "../review";
import {
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const apply = (document: Document, op: DocumentOp): Document => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const resolve = (
  document: Document,
  ids: readonly number[],
  decision: RevisionDecision,
): Document =>
  apply(document, {
    type: "resolveRevision",
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision,
  });

const link = (children: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children,
});

const links = [link([]), link([{ type: "run", content: [{ type: "tab" }] }])];

/** Every tracked record's id, pieces included. */
const wrapperIds = (document: Document): number[] => {
  const ids = new Set<number>();
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node)) ids.add(node.info.id);
      visit(childNodes(node) ?? []);
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") visit(block.content);
  return [...ids];
};

const cases = [
  { type: "insertion", decision: REVISION_DECISIONS.ACCEPT },
  { type: "deletion", decision: REVISION_DECISIONS.REJECT },
  { type: "moveFrom", decision: REVISION_DECISIONS.REJECT },
] as const;

for (const { type, decision } of cases)
  test(`a split between two links in a ${type} rejoins them apart once the ${type} is ${decision}ed`, () => {
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [{ type, info: { id: 1, author: "A" }, content: links }],
            },
          ],
        },
      },
    };
    const split = apply(source, {
      type: "splitBlock",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      newBlockId: "10000000",
      newIds: { revision: [700, 701, 702], control: [800] },
      revision: { id: 2, author: "C", date: "2026-03-04T05:00:00Z" },
    });
    const pieces = wrapperIds(split);
    const results = [
      resolve(resolve(split, pieces, decision), [2], REVISION_DECISIONS.REJECT),
      resolve(resolve(split, [2], REVISION_DECISIONS.REJECT), pieces, decision),
    ];
    if (decision === REVISION_DECISIONS.REJECT)
      results.push(resolve(split, [...pieces, 2], decision));
    for (const resolved of results)
      expect(resolved.package.document.content).toMatchObject([
        { type: "paragraph", content: links },
      ]);
  });
