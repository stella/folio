import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

// Tracked deletions of a container's last paragraphs, one at a time in any
// order, resolve as the same deletions made directly: accepting leaves the
// paragraphs before them, rejecting gives the document back.

const paragraph = (index: number, text: string): Paragraph => ({
  type: "paragraph",
  paraId: (index + 1).toString(16).padStart(8, "0"),
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (texts: readonly string[]): Document => ({
  package: {
    document: { content: texts.map((text, index) => paragraph(index, text)) },
  },
});

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const permutations = (items: readonly number[]): number[][] =>
  items.length <= 1
    ? [[...items]]
    : items.flatMap((item, index) =>
        permutations(items.filter((_, other) => other !== index)).map((rest) =>
          [item].concat(rest),
        ),
      );

const idOf = (index: number) => (index + 1).toString(16).padStart(8, "0");

const words = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((node) =>
            node.type === "run"
              ? node.content.flatMap((leaf) =>
                  leaf.type === "text" ? [leaf.text] : [],
                )
              : [],
          )
          .join("")
      : block.type,
  );

for (const texts of [
  ["one", "two", "three", "four"],
  ["one", "two", "three", ""],
] as const)
  for (const count of texts.at(-1) === "" ? [2, 3] : [2, 3, 4])
    for (const order of permutations(
      Array.from({ length: count }, (_, index) => texts.length - count + index),
    ))
      test(`deleting the last ${count} of ${JSON.stringify(texts)} in order ${order.join(",")} resolves as direct`, () => {
        const original = documentOf(texts);
        let tracked = original;
        const revisions: number[] = [];
        for (const [step, index] of order.entries()) {
          const id = 100 + step * 10;
          const deleted = apply(tracked, {
            type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
            story: OP_STORIES.MAIN,
            blockIds: [idOf(index)],
            revision: { id, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
            newIds: { revision: [id + 1, id + 2, id + 3, id + 4] },
          });
          revisions.push(...deleted.revisions);
          tracked = deleted.document;
        }
        const resolve = (decision: RevisionDecision) =>
          apply(tracked, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: revisions,
            decision,
          }).document;
        // Deleted directly, the paragraphs before them are left; deleting
        // every paragraph leaves the first, emptied, as the container's end.
        expect(words(resolve(REVISION_DECISIONS.ACCEPT))).toEqual(
          count === texts.length ? [""] : texts.slice(0, texts.length - count),
        );
        expect(words(resolve(REVISION_DECISIONS.REJECT))).toEqual(
          words(original),
        );
      });
