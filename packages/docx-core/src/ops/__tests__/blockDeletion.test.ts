import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DeleteBlocksOp,
  type DocumentOp,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const documentOf = (content: BlockContent[]): Document => ({ package: { document: { content } } });
const cell = (content: BlockContent[]): BlockContent => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
});
const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};
const deletion = (blockIds: string[]): DeleteBlocksOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  story: OP_STORIES.MAIN,
  blockIds,
  revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
  newIds: { revision: Array.from({ length: 32 }, (_, index) => 101 + index) },
});
const assertLaws = (document: Document, op: DeleteBlocksOp) => {
  const tracked = apply(document, op);
  const direct = apply(document, { ...op, revision: undefined });
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const resolved = apply(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.revisions,
      decision,
    });
    expect(resolved.document).toStrictEqual(
      decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
    );
    const undo = applyDocumentOps(resolved.document, resolved.inverse);
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(tracked.document);
  }
  for (const edit of [tracked, direct]) {
    const undo = applyDocumentOps(edit.document, edit.inverse);
    if (undo.isErr()) throw undo.error;
    expect(undo.value.document).toStrictEqual(document);
  }
  return { tracked, direct };
};

test.each(["body", "cell"] as const)(
  "whole-container deletion retains the last paragraph in a %s",
  (container) => {
    const first = {
      ...paragraph("00000001", "first"),
      formatting: { alignment: "start" },
    } satisfies Paragraph;
    const final = {
      ...paragraph("00000002", "last"),
      formatting: { alignment: "end", runProperties: { bold: true } },
    } satisfies Paragraph;
    const content = [first, final];
    const document = documentOf(container === "body" ? content : [cell(content)]);
    const { tracked, direct } = assertLaws(document, deletion(["00000001", "00000002"]));
    expect(storyParagraphs(direct.document.package.document).map(({ paragraph: p }) => p)).toEqual([
      { ...final, content: [] },
    ]);
    expect(
      storyParagraphs(tracked.document.package.document).at(-1)?.paragraph.pPrMark,
    ).toBeUndefined();
  },
);

test.each(["body", "cell"] as const)(
  "terminal deletion compensates properties and retains the final identity in a %s",
  (container) => {
    for (const text of ["", "final"]) {
      const previous = {
        ...paragraph("00000001", "kept"),
        formatting: { alignment: "center", runProperties: { italic: true } },
      } satisfies Paragraph;
      const final = {
        ...paragraph("00000002", text),
        formatting: { alignment: "end", runProperties: { bold: true } },
      } satisfies Paragraph;
      const content = [previous, final];
      const document = documentOf(container === "body" ? content : [cell(content)]);
      const { direct } = assertLaws(document, deletion(["00000002"]));
      expect(storyParagraphs(direct.document.package.document).at(0)?.paragraph).toEqual({
        ...final,
        content: previous.content,
        formatting: { alignment: "center", runProperties: { bold: true } },
      });
    }
  },
);

test("interior deletion preserves the following object and reports physical changes", () => {
  const first = paragraph("00000001", "remove");
  const following = paragraph("00000002", "kept");
  const { direct, tracked } = assertLaws(documentOf([first, following]), deletion(["00000001"]));
  expect(direct.document.package.document.content.at(0)).toBe(following);
  expect(tracked.document.package.document.content.at(-1)).toBe(following);
  expect(direct.touched).toEqual({ modified: [], inserted: [], removed: [first.paraId] });
  expect(tracked.touched).toEqual({ modified: [first.paraId], inserted: [], removed: [] });
});

test("a sole paragraph after a table is cleared without marking its final break", () => {
  const table = cell([paragraph("00000003", "cell")]);
  const final = paragraph("00000001", "remove");
  const { direct } = assertLaws(documentOf([table, final]), deletion(["00000001"]));
  expect(direct.document.package.document.content).toEqual([table, { ...final, content: [] }]);
  expect(direct.document.package.document.content.at(0)).toBe(table);
});

test("deletion refuses invalid selections, boundaries and paragraph review conflicts", () => {
  const first = paragraph("00000001", "first");
  const final = paragraph("00000002", "final");
  const marked = {
    ...first,
    pPrMark: { kind: "ins", info: { id: 9, author: "A" } },
  } satisfies Paragraph;
  const cases = [
    {
      document: documentOf([first, final]),
      ids: [],
      reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
    },
    {
      document: documentOf([first, final]),
      ids: ["00000002", "00000001"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    },
    {
      document: documentOf([first, cell([final])]),
      ids: ["00000001", "00000002"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    },
    {
      document: documentOf([first, cell([final])]),
      ids: ["00000001"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    },
    {
      document: documentOf([marked, final]),
      ids: ["00000001"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    },
    {
      document: documentOf([{ ...first, sectionProperties: {} }, final]),
      ids: ["00000001"],
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    },
  ];
  for (const { document, ids, reason } of cases) {
    const result = applyDocumentOp(document, deletion(ids));
    expect(result.isErr() && result.error.reason).toBe(reason);
  }
});

test("tracked deletion refuses existing inline review and missing physical ids", () => {
  const first = paragraph("00000001", "first");
  const final = paragraph("00000002", "final");
  const reviewed: Paragraph = {
    ...first,
    content: [
      {
        type: "deletion",
        info: { id: 9, author: "A" },
        content: [{ type: "run", content: [{ type: "text", text: "first" }] }],
      },
    ],
  };
  const refused = applyDocumentOp(documentOf([reviewed, final]), deletion(["00000001"]));
  expect(refused.isErr() && refused.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  const missing = applyDocumentOp(documentOf([first, final]), {
    ...deletion(["00000001"]),
    newIds: {},
  });
  expect(missing.isErr() && missing.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS);
});
