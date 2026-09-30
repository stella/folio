import { expect, test } from "bun:test";

import {
  paragraphNumberingReference,
  type BlockContent,
  type Document,
  type Paragraph,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import { revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DocumentOp,
  type InsertBlocksOp,
  type RevisionDecision,
} from "../types";

const makeParagraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const baseDocument: Document = {
  package: { document: { content: [makeParagraph("00000001", "kept")] } },
};

test("inserting paragraphs before an anchor preserves its identity and content", () => {
  const inserted = makeParagraph("00000002", "new");
  const result = applyDocumentOp(baseDocument, {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: "00000001" },
    blocks: [inserted],
  });
  if (result.isErr()) throw result.error;
  expect(result.value.document.package.document.content).toEqual([
    inserted,
    ...baseDocument.package.document.content,
  ]);
  expect(result.value.document.package.document.content.at(-1)).toBe(
    baseDocument.package.document.content.at(0),
  );
});

const stamp = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (
  applied: Pick<AppliedDocumentOp, "document" | "revisions">,
  decision: RevisionDecision,
) =>
  apply(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: applied.revisions,
    decision,
  });

const undo = (applied: AppliedDocumentOp) => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value.document;
};

const documentOf = (content: BlockContent[]): Document => ({ package: { document: { content } } });

const inCell = (content: BlockContent[]): BlockContent => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
});

const insertion = (blocks: Paragraph[]): InsertBlocksOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "before", blockId: "00000001" },
  blocks,
  revision: stamp,
  newIds: { revision: Array.from({ length: 32 }, (_, index) => 101 + index) },
});

test.each(["body", "cell"] as const)(
  "tracked list insertion in a %s accepts without a leftover paragraph",
  (container) => {
    const original = makeParagraph("00000001", "kept");
    const content = [original];
    const document = documentOf(container === "body" ? content : [inCell(content)]);
    const listed = [makeParagraph("00000002", "first"), makeParagraph("00000003", "second")].map(
      (paragraph) =>
        Object.assign({}, paragraph, {
          formatting: { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) },
        }),
    );
    const op = insertion(listed);
    const tracked = apply(document, op);
    const direct = apply(document, { ...op, revision: undefined });
    expect(resolve(tracked, REVISION_DECISIONS.ACCEPT).document).toEqual(direct.document);
    expect(resolve(tracked, REVISION_DECISIONS.REJECT).document).toEqual(document);
    expect(undo(tracked)).toEqual(document);
    const demand = revisionIdDemand(document, op);
    if (demand.isErr()) throw demand.error;
    expect(demand.value).toBe(3);
  },
);

test.each(["", "kept"])(
  "tracked insertion before a final paragraph preserves its formatting and id (%s)",
  (text) => {
    const final = {
      ...makeParagraph("00000001", text),
      formatting: { alignment: "end" },
    } satisfies Paragraph;
    const document = documentOf([inCell([final]), makeParagraph("00000004", "outside")]);
    const tracked = apply(document, insertion([makeParagraph("00000002", "new")]));
    const finalAfter = storyParagraphs(tracked.document.package.document).find(
      ({ paragraph }) => paragraph.paraId === final.paraId,
    );
    expect(finalAfter?.paragraph).toBe(final);
    expect(finalAfter?.paragraph.pPrMark).toBeUndefined();
    expect(resolve(tracked, REVISION_DECISIONS.REJECT).document).toEqual(document);
  },
);

test("direct insertion after a final paragraph is exactly reversible", () => {
  const op = {
    ...insertion([makeParagraph("00000002", "new")]),
    at: { type: "after", blockId: "00000001" },
    revision: undefined,
  } as const;
  const inserted = apply(baseDocument, op);
  expect(inserted.document.package.document.content).toEqual([
    ...baseDocument.package.document.content,
    ...op.blocks,
  ]);
  expect(undo(inserted)).toEqual(baseDocument);
  expect(inserted.touched).toEqual({ modified: [], inserted: ["00000002"], removed: [] });
});

test("tracked insertion refuses terminal positions and non-paragraph boundaries", () => {
  const op = {
    ...insertion([makeParagraph("00000002", "new")]),
    at: { type: "after", blockId: "00000001" },
  } as const;
  for (const document of [
    documentOf([makeParagraph("00000001", "kept")]),
    documentOf([makeParagraph("00000001", "kept"), inCell([makeParagraph("00000003", "cell")])]),
  ]) {
    const result = applyDocumentOp(document, op);
    expect(result.isErr() && result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  }
});

test("insertion validates identities, review conflicts and section boundaries", () => {
  const collision = applyDocumentOp(
    baseDocument,
    insertion([makeParagraph("00000001", "duplicate")]),
  );
  expect(collision.isErr() && collision.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
  );
  const changed = {
    ...makeParagraph("00000002", "changed"),
    pPrMark: { kind: "ins", info: { id: 9, author: "A" } },
  } satisfies Paragraph;
  const conflict = applyDocumentOp(baseDocument, insertion([changed]));
  expect(conflict.isErr() && conflict.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  const section = { ...makeParagraph("00000002", "section"), sectionProperties: {} };
  const boundary = applyDocumentOp(baseDocument, insertion([section]));
  expect(boundary.isErr() && boundary.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  const insufficient = applyDocumentOp(baseDocument, {
    ...insertion([makeParagraph("00000002", "new")]),
    newIds: {},
  });
  expect(insufficient.isErr() && insufficient.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );
});
