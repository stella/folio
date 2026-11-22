/**
 * A deletion inside a container and a pending insertion of a marker right
 * after it, which split the container in two. Accepting the deletion empties
 * the first piece away; rejecting the insertion later keeps the marker in
 * the piece left across from it, as rejecting it first, while the pieces
 * still rejoin around it, does.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const wrapper = (content: ParagraphContent[]): ParagraphContent => ({
  type: "inlineWrapper",
  kind: "bidi",
  control: "embedding",
  content,
});

const marker: ParagraphContent = {
  type: "moveToRangeStart",
  id: 1,
  name: "move1",
  author: "A",
};

const document: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            text("a"),
            wrapper([
              {
                type: "deletion",
                info: { id: 1, author: "C", date: "2026-03-04T05:00:00Z" },
                content: [text("a")],
                resolutionJoins: { before: 0, after: 2, remove: 0 },
              },
            ]),
            {
              type: "insertion",
              info: { id: 2, author: "C", date: "2026-03-04T05:00:01Z" },
              content: [marker],
              resolutionJoins: { before: 0, after: 0, remove: 1 },
            },
            wrapper([text("aaa")]),
          ],
        },
      ],
    },
  },
};

const resolved = (
  from: Document,
  id: number,
  decision: RevisionDecision,
): Document => {
  const applied = applyDocumentOp(from, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [id],
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const contentOf = (from: Document) => {
  const [paragraph] = from.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : undefined;
};

test("a marker kept beside an emptied piece goes into the piece across in either order", () => {
  const { ACCEPT, REJECT } = REVISION_DECISIONS;
  const acceptedFirst = contentOf(
    resolved(resolved(document, 1, ACCEPT), 2, REJECT),
  );
  const rejectedFirst = contentOf(
    resolved(resolved(document, 2, REJECT), 1, ACCEPT),
  );
  expect(acceptedFirst).toEqual([text("a"), wrapper([marker, text("aaa")])]);
  expect(acceptedFirst).toEqual(rejectedFirst);
});
