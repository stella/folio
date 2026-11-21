/**
 * Deleting, tracked or suggested, the paragraphs one's own pending insertion
 * added takes that insertion back: the document is the one before it, ids
 * and paragraph properties included, after a paragraph another follows and
 * after the container's last paragraph, whose former mark the inserted
 * paragraph took over. No deletion of an addition is recorded; a comment it
 * leaves nothing of retires with it, as a direct deletion's does. A deletion
 * by another suggestion is recorded beside the insertion instead.
 */
import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphFormatting,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  type AppliedDocumentOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionStamp,
} from "../types";

const paragraph = (
  paraId: string,
  text: string,
  formatting?: ParagraphFormatting,
): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
  ...(formatting === undefined ? {} : { formatting }),
});

const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const apply = (document: Document, op: DocumentOp): AppliedDocumentOp => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const undo = (applied: AppliedDocumentOp): Document => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value.document;
};

const stampOf = (id: number, suggestionId?: string): RevisionStamp => ({
  id,
  author: "Reviewer",
  date: "2026-02-03T04:05:06Z",
  ...(suggestionId === undefined ? {} : { suggestionId }),
});

const newIds = (from: number) => ({
  revision: Array.from({ length: 16 }, (_, index) => from + index),
});

const INSERTED_ID = "0000000A";
const STAMPS = {
  tracked: undefined,
  suggested: "s1",
} as const;
const PLACES = {
  // The paragraph the insertion follows, and the document it goes into.
  "after a paragraph another follows": "00000001",
  "after the container's last paragraph": "00000002",
} as const;
const FORMATTINGS = {
  "the anchor's properties": undefined,
  "properties of its own": { alignment: "center" },
} as const satisfies Record<string, ParagraphFormatting | undefined>;
const TEXTS = { words: "Inserted.", "no words": "" } as const;

const original = documentOf([
  paragraph("00000001", "First.", { alignment: "start" }),
  paragraph("00000002", "Last.", { alignment: "start" }),
]);

type InsertedOptions = {
  anchor: string;
  text?: string;
  formatting?: ParagraphFormatting | undefined;
  suggestionId: string | undefined;
};

const inserted = ({
  anchor,
  text = TEXTS.words,
  formatting = { alignment: "start" },
  suggestionId,
}: InsertedOptions) =>
  apply(original, {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: anchor },
    blocks: [paragraph(INSERTED_ID, text, formatting)],
    revision: stampOf(10, suggestionId),
    newIds: newIds(11),
  });

const deleted = (document: Document, suggestionId: string | undefined) =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: [INSERTED_ID],
    revision: stampOf(30, suggestionId),
    newIds: newIds(31),
  });

for (const [stampName, suggestionId] of Object.entries(STAMPS))
  for (const [place, anchor] of Object.entries(PLACES))
    for (const [formattingName, formatting] of Object.entries(FORMATTINGS))
      for (const [textName, text] of Object.entries(TEXTS))
        test(`deleting its own ${stampName} insertion of ${textName} ${place}, with ${formattingName}, takes it back`, () => {
          const insertion = inserted({
            anchor,
            text,
            formatting,
            suggestionId,
          });
          const [, second, last] = insertion.document.package.document.content;
          // After the last paragraph, the anchor gets the new break and the
          // inserted paragraph its former mark; elsewhere it adds its own.
          if (anchor === PLACES["after the container's last paragraph"])
            expect(
              second?.type === "paragraph" && second.pPrMark?.survivor,
            ).toBe("first");
          else
            expect(second?.type === "paragraph" && second.pPrMark?.kind).toBe(
              "ins",
            );
          expect(last?.type === "paragraph" && last.paraId).toBe(
            anchor === PLACES["after the container's last paragraph"]
              ? INSERTED_ID
              : "00000002",
          );
          const deletion = deleted(insertion.document, suggestionId);
          expect(deletion.document).toStrictEqual(original);
          expect(undo(deletion)).toStrictEqual(insertion.document);
        });

for (const [place, anchor] of Object.entries(PLACES))
  test(`a later suggestion's deletion ${place} is recorded beside the insertion`, () => {
    const insertion = inserted({ anchor, suggestionId: "s1" });
    const deletion = deleted(insertion.document, "s2");
    expect(deletion.document).not.toStrictEqual(original);
    expect(deletion.revisions.length).toBeGreaterThan(0);
    // Rejecting the deletion alone gives the insertion back.
    const rejected = apply(deletion.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: deletion.revisions,
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(rejected.document.package.document.content).toHaveLength(3);
    // Rejecting both leaves the document before them.
    const both = apply(rejected.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: insertion.revisions,
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(both.document).toStrictEqual(original);
  });

type CommentedOptions = {
  document: Document;
  from: { blockId: string; offset: number };
  to: { blockId: string; offset: number };
};

const commented = ({ document, from, to }: CommentedOptions) =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 5,
      author: "Reader",
      content: [paragraph("0000000C", "Why?")],
    },
    anchor: {
      kind: "range",
      from: { story: OP_STORIES.MAIN, ...from },
      to: { story: OP_STORIES.MAIN, ...to },
    },
  });

for (const [stampName, suggestionId] of Object.entries(STAMPS))
  for (const [place, anchor] of Object.entries(PLACES))
    test(`deleting its own ${stampName} insertion ${place} retires a comment held in it`, () => {
      const insertion = inserted({ anchor, suggestionId });
      const withComment = commented({
        document: insertion.document,
        from: { blockId: INSERTED_ID, offset: 0 },
        to: { blockId: INSERTED_ID, offset: TEXTS.words.length },
      });
      const deletion = deleted(withComment.document, suggestionId);
      // The comment part creating the comment added stays, empty.
      expect(deletion.document.package.document.content).toStrictEqual(
        original.package.document.content,
      );
      expect(deletion.document.package.document.comments ?? []).toEqual([]);
      expect(undo(deletion)).toStrictEqual(withComment.document);
    });

test("a comment reaching past its own insertion keeps the insertion's words", () => {
  const insertion = inserted({
    anchor: PLACES["after a paragraph another follows"],
    suggestionId: undefined,
  });
  const withComment = commented({
    document: insertion.document,
    from: { blockId: INSERTED_ID, offset: 0 },
    to: { blockId: "00000002", offset: 2 },
  });
  const deletion = deleted(withComment.document, undefined);
  expect(deletion.document.package.document.comments).toHaveLength(1);
  expect(undo(deletion)).toStrictEqual(withComment.document);
});
