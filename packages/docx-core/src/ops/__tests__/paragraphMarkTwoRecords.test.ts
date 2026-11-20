/**
 * A paragraph mark holding another author's insertion and a later deletion
 * of it: each record resolves on its own, in either order, and together.
 */
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const INSERTION = 1;
const DELETION = 2;

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const document: Document = {
  package: {
    document: {
      content: [
        {
          ...paragraph("00000001", "One"),
          pPrMark: {
            kind: "ins",
            info: { id: INSERTION, author: "A", date: "2026-01-01T00:00:00Z" },
            deletion: {
              info: { id: DELETION, author: "B", date: "2026-01-02T00:00:00Z" },
            },
          },
        },
        paragraph("00000002", "Two"),
      ],
    },
  },
};

const resolved = (
  source: Document,
  revisionIds: number[],
  decision: RevisionDecision,
): Document => {
  const applied = applyDocumentOp(source, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (applied.isErr()) throw applied.error;
  const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(source);
  return applied.value.document;
};

const shape = (result: Document) =>
  result.package.document.content.map((block) =>
    block.type === "paragraph"
      ? {
          text: block.content
            .flatMap((item) => (item.type === "run" ? item.content : []))
            .map((item) => (item.type === "text" ? item.text : ""))
            .join(""),
          mark: block.pPrMark && {
            kind: block.pPrMark.kind,
            id: block.pPrMark.info.id,
            ...(block.pPrMark.deletion
              ? { deletion: block.pPrMark.deletion.info.id }
              : {}),
          },
        }
      : block.type,
  );

const JOINED = [{ text: "OneTwo", mark: undefined }];

test("accepting the insertion alone leaves the deletion pending", () => {
  expect(
    shape(resolved(document, [INSERTION], REVISION_DECISIONS.ACCEPT)),
  ).toEqual([
    { text: "One", mark: { kind: "del", id: DELETION } },
    { text: "Two", mark: undefined },
  ]);
});

test("rejecting the deletion alone leaves the insertion pending", () => {
  expect(
    shape(resolved(document, [DELETION], REVISION_DECISIONS.REJECT)),
  ).toEqual([
    { text: "One", mark: { kind: "ins", id: INSERTION } },
    { text: "Two", mark: undefined },
  ]);
});

for (const [ids, decision] of [
  [[DELETION], REVISION_DECISIONS.ACCEPT],
  [[INSERTION], REVISION_DECISIONS.REJECT],
  [[INSERTION, DELETION], REVISION_DECISIONS.ACCEPT],
  [[INSERTION, DELETION], REVISION_DECISIONS.REJECT],
] as const) {
  test(`${decision}ing ${JSON.stringify(ids)} takes the break`, () => {
    expect(shape(resolved(document, [...ids], decision))).toEqual(JOINED);
  });
}

test("each order of accepting the insertion and the deletion joins", () => {
  const insertionFirst = resolved(
    resolved(document, [INSERTION], REVISION_DECISIONS.ACCEPT),
    [DELETION],
    REVISION_DECISIONS.ACCEPT,
  );
  expect(shape(insertionFirst)).toEqual(JOINED);
});

test("accepting the insertion then rejecting the deletion keeps the break", () => {
  expect(
    shape(
      resolved(
        resolved(document, [INSERTION], REVISION_DECISIONS.ACCEPT),
        [DELETION],
        REVISION_DECISIONS.REJECT,
      ),
    ),
  ).toEqual([
    { text: "One", mark: undefined },
    { text: "Two", mark: undefined },
  ]);
  expect(
    shape(
      resolved(
        resolved(document, [DELETION], REVISION_DECISIONS.REJECT),
        [INSERTION],
        REVISION_DECISIONS.ACCEPT,
      ),
    ),
  ).toEqual([
    { text: "One", mark: undefined },
    { text: "Two", mark: undefined },
  ]);
});
