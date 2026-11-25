/**
 * Accepting an insertion can retire one deferred seam and leave another,
 * shallower one waiting on the insertion after it: the kept markers between
 * the two insertions keep the name they wait with.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const link = (value: string): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children: [text(value)],
});

const info = (id: number, second: number) => ({
  id,
  author: "A",
  date: `2026-03-04T05:00:0${second}Z`,
});

const KEPT = {
  keptFrom: "2026-03-04T05:00:01Z",
  holders: { before: 2, after: 3, scope: "seam" },
} as const;

const documentOf = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            { type: "bookmarkStart", id: 1, name: "_Ref1" },
            link("a"),
            {
              type: "insertion",
              info: info(2, 3),
              content: [link("x")],
              resolutionJoins: {
                before: 1,
                after: 1,
                remove: 0,
                deferredRemove: [
                  { depth: 1, blockers: [2, 3] },
                  { depth: 2, blockers: [2] },
                ],
              },
            },
            { type: "bookmarkEnd", id: 1, keptCut: KEPT },
            {
              type: "insertion",
              info: info(3, 4),
              content: [link("y")],
              resolutionJoins: {
                before: 0,
                after: 0,
                remove: 0,
                deferredRemove: [{ depth: 1, blockers: [2, 3] }],
              },
            },
            link("b"),
          ],
        },
      ],
    },
  },
});

const resolved = (
  document: Document,
  revisionIds: number[],
  decision: "accept" | "reject",
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision:
      decision === "accept"
        ? REVISION_DECISIONS.ACCEPT
        : REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const contentOf = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

test("a marker between an accepted blocker and one a seam still waits on keeps its name", () => {
  const content = contentOf(resolved(documentOf(), [2], "accept"));
  expect(content).toContainEqual({
    type: "bookmarkEnd",
    id: 1,
    keptCut: { keptFrom: KEPT.keptFrom, holders: { after: 3, scope: "seam" } },
  });
});
