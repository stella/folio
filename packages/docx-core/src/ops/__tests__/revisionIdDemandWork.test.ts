import { expect, spyOn, test } from "bun:test";

import {
  type Document,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { validateOpsDocument } from "../contract";
import { formatParaId } from "../idLease";
import * as ids from "../ids";
import { revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  type InsertContentOp,
  OP_STORIES,
} from "../types";

const paragraph = (id: number, content: Paragraph["content"]): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content,
});
const run = (text: string, revisionId?: number): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
  ...(revisionId === undefined
    ? {}
    : {
        propertyChanges: [
          {
            type: "runPropertyChange",
            info: { id: revisionId, author: "Other" },
          },
        ],
      }),
});
const control = (id: number): ParagraphContent => ({
  type: "inlineSdt",
  properties: { sdtType: "richText", id },
  content: [run("control")],
});
const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});
const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: formatParaId(1),
  offset,
});
const revision = { id: 99, author: "Author", date: "2026-01-11T00:00:00.000Z" };
const cut: DeleteRangeOp = {
  type: DOCUMENT_OP_TYPES.DELETE_RANGE,
  from: at(1),
  to: at(4),
  revision,
};

test("ordinary range demand does not flatten package IDs for a successful no-extra-ID deletion", () => {
  const doc = documentOf([
    paragraph(1, [run("abcde")]),
    ...Array.from({ length: 127 }, (_, index) =>
      paragraph(index + 2, [run("untouched", index + 1000)]),
    ),
    paragraph(129, [run("highest", MAX_REVISION_ID), control(MAX_REVISION_ID)]),
  ]);
  expect(validateOpsDocument(doc).isOk()).toBe(true);
  const flatten = spyOn(ids, "packageIdentityKeys");
  try {
    const invalidUnusedPool = {
      control: [-1],
      revision: [MAX_REVISION_ID + 1],
    };
    const direct: DeleteRangeOp = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: cut.from,
      to: cut.to,
      newIds: invalidUnusedPool,
    };
    expect(revisionIdDemand(doc, direct).unwrap()).toBe(0);
    expect(
      revisionIdDemand(doc, { ...cut, newIds: invalidUnusedPool }).unwrap(),
    ).toBe(0);
    expect(
      flatten.mock.calls.filter(([pkg]) => pkg === doc.package),
    ).toHaveLength(0);
  } finally {
    flatten.mockRestore();
  }
});

test("a real ID-consuming cut retains the original demand and exhaustion refusal", () => {
  const anchor = paragraph(1, [run("abc", 4), run("de")]);
  const available = documentOf([anchor]);
  const empty = applyDocumentOp(available, {
    ...cut,
    newIds: { revision: [], control: [] },
  });
  expect(empty.isErr() && empty.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );
  expect(revisionIdDemand(available, cut).unwrap()).toBe(1);
  const exhausted = documentOf([
    anchor,
    paragraph(2, [run("max", MAX_REVISION_ID)]),
  ]);
  const refusal = revisionIdDemand(exhausted, cut);
  expect(refusal.isErr() && refusal.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
  );
});

test("a different operation still receives speculative controls and refuses exhausted consumed IDs", () => {
  const anchor = paragraph(1, [control(3)]);
  const operation: InsertContentOp = {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at(0),
    slice: { content: [control(3)], openStart: 0, openEnd: 0 },
  };
  const available = documentOf([anchor]);
  const empty = applyDocumentOp(available, {
    ...operation,
    newIds: { control: [] },
  });
  expect(empty.isErr() && empty.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );
  expect(revisionIdDemand(available, operation).unwrap()).toBe(0);
  const exhausted = documentOf([
    anchor,
    paragraph(2, [control(MAX_REVISION_ID)]),
  ]);
  const refusal = revisionIdDemand(exhausted, operation);
  expect(refusal.isErr() && refusal.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
  );
});

test("presence metadata retains the original sampled pool and refusal", () => {
  const doc = documentOf([paragraph(1, [run("abcde")])]);
  const counts = spyOn(ids, "packageIdentityKeys");
  try {
    expect(validateOpsDocument(doc).isOk()).toBe(true);
    counts.mockClear();
    expect(
      revisionIdDemand(doc, {
        ...cut,
        undefinedFields: [["expected"]],
      }).unwrap(),
    ).toBe(0);
    expect(
      counts.mock.calls.filter(([pkg]) => pkg === doc.package),
    ).toHaveLength(2);
    counts.mockClear();

    const refused = revisionIdDemand(doc, {
      ...cut,
      undefinedFields: [["newIds", "control", "0"]],
    });
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    expect(refused.isErr() && refused.error.message).toBe(
      "Operation presence metadata conflicts with an existing value.",
    );
    expect(
      counts.mock.calls.filter(([pkg]) => pkg === doc.package),
    ).toHaveLength(2);
  } finally {
    counts.mockRestore();
  }
});
