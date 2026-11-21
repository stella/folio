import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { formatParaId } from "../idLease";
import { childNodes, type InlineNode } from "../leaves";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
  type RevisionStamp,
} from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (id: number, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content,
});
const reference: ParagraphContent = {
  type: "run",
  content: [{ type: "footnoteRef", id: 1 }],
};

/** A body citing footnote 1, a second time when `shared`, and the two-paragraph note. */
const documentOf = (shared: boolean): Document => ({
  package: {
    document: {
      content: [
        paragraph(1, [run("Noted"), reference, run(" text.")]),
        paragraph(2, shared ? [run("Again"), reference] : [run("Tail.")]),
      ],
    },
    footnotes: [
      {
        type: "footnote",
        id: 1,
        content: [
          paragraph(101, [run("First note line.")]),
          paragraph(102, [run("Second note line.")]),
        ],
      },
    ],
  },
});

const STAMP = {
  id: 41,
  author: "Reviewer",
  date: "2026-11-21T00:00:00Z",
} as const satisfies RevisionStamp;
const NEW_IDS = Array.from({ length: 16 }, (_, index) => 42 + index);

const deleteReference = (
  revision: RevisionStamp,
  newIds: readonly number[] = NEW_IDS,
): DocumentOp => {
  const at = { story: OP_STORIES.MAIN, blockId: formatParaId(1) } as const;
  return {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: { ...at, offset: 5 },
    to: { ...at, offset: 6 },
    revision,
    newIds: { revision: [...newIds] },
  };
};

const noteParagraphs = (document: Document) =>
  storyParagraphs(document.package.footnotes?.at(0) ?? { content: [] }).map(
    ({ paragraph: found }) => found,
  );

/** Every deletion record in the note's text, and whether any text stands outside one. */
const noteDeletions = (document: Document) => {
  const deletions: { id: number; author: string; date?: string }[] = [];
  let live = false;
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (node.type === "deletion") {
        deletions.push(node.info);
        continue;
      }
      if (node.type === "text") live = true;
      visit(childNodes(node) ?? []);
    }
  };
  for (const found of noteParagraphs(document)) visit(found.content);
  return { deletions, live };
};

const resolve = (document: Document, decision: RevisionDecision) =>
  applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [STAMP.id],
      decision,
    },
  ]).unwrap().document;

describe("a tracked deletion of a note's last live reference", () => {
  test("deletes the note's text and marks under fresh ids of the operation", () => {
    const document = documentOf(false);
    const edit = applyDocumentOps(document, [deleteReference(STAMP)]).unwrap();
    expect(contractViolation(edit.document)).toBeUndefined();
    const { deletions, live } = noteDeletions(edit.document);
    expect(live).toBe(false);
    const marks = noteParagraphs(edit.document).map(({ pPrMark }) => pPrMark);
    expect(marks.map((mark) => mark?.kind)).toEqual(["del", "del"]);
    const records = [
      ...deletions,
      ...marks.flatMap((mark) => (mark === undefined ? [] : [mark.info])),
    ];
    expect(records).toHaveLength(4);
    for (const { author, date } of records)
      expect({ author, date }).toEqual({
        author: STAMP.author,
        date: STAMP.date,
      });
    const ids = records.map(({ id }) => id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const id of ids) {
      expect(NEW_IDS).toContain(id);
      expect(edit.revisions).toContain(id);
    }
    // Undo is exact.
    expect(
      applyDocumentOps(edit.document, edit.inverse).unwrap().document,
    ).toStrictEqual(document);
  });

  test("rejecting the reference's deletion gives the note its text back", () => {
    const document = documentOf(false);
    const pending = applyDocumentOps(document, [
      deleteReference(STAMP),
    ]).unwrap().document;
    expect(resolve(pending, REVISION_DECISIONS.REJECT)).toStrictEqual(document);
  });

  test("accepting the reference's deletion removes the note", () => {
    const pending = applyDocumentOps(documentOf(false), [
      deleteReference(STAMP),
    ]).unwrap().document;
    expect(
      resolve(pending, REVISION_DECISIONS.ACCEPT).package.footnotes,
    ).toEqual([]);
  });

  test("refuses without the ids the note's records need", () => {
    const refused = applyDocumentOps(documentOf(false), [
      deleteReference(STAMP, []),
    ]);
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
  });

  test("leaves a note another live reference still cites", () => {
    const document = documentOf(true);
    const edit = applyDocumentOps(document, [deleteReference(STAMP)]).unwrap();
    expect(edit.document.package.footnotes).toBe(document.package.footnotes);
  });

  test("leaves the note of a suggestion, which stays its author's until accepted", () => {
    const document = documentOf(false);
    const edit = applyDocumentOps(document, [
      deleteReference({ ...STAMP, suggestionId: "suggestion-1" }),
    ]).unwrap();
    expect(edit.document.package.footnotes).toBe(document.package.footnotes);
  });
});
