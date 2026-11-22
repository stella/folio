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
import { noteDeletionOps } from "../noteDeletion";
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

const deleteReference = (revision: RevisionStamp): DocumentOp => {
  const at = { story: OP_STORIES.MAIN, blockId: formatParaId(1) } as const;
  return {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: { ...at, offset: 5 },
    to: { ...at, offset: 6 },
    revision,
  };
};

/** The reference's tracked deletion, then the operations its note follows it with. */
const deleteWithNote = (
  document: Document,
  stamp: RevisionStamp,
  revisionIds: readonly number[] = NEW_IDS,
) => {
  const deletion = applyDocumentOps(document, [
    deleteReference(stamp),
  ]).unwrap();
  return {
    deletion,
    noteOps: noteDeletionOps({
      before: document,
      after: deletion.document,
      stamp,
      revisionIds,
    }),
  };
};

const followed = (document: Document, stamp: RevisionStamp = STAMP) => {
  const { noteOps } = deleteWithNote(document, stamp);
  return applyDocumentOps(document, [
    deleteReference(stamp),
    ...noteOps.unwrap().ops,
  ]).unwrap();
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
  test("leaves the note alone: the operation touches only what it names", () => {
    const document = documentOf(false);
    const { deletion } = deleteWithNote(document, STAMP);
    expect(deletion.document.package.footnotes).toBe(
      document.package.footnotes,
    );
  });

  test("is followed by the note's text and marks deleted under fresh ids", () => {
    const document = documentOf(false);
    const edit = followed(document);
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
    const { noteOps } = deleteWithNote(document, STAMP);
    for (const id of ids) expect(NEW_IDS).toContain(id);
    expect(new Set(noteOps.unwrap().revisions)).toEqual(new Set(ids));
    // Undo is exact.
    expect(
      applyDocumentOps(edit.document, edit.inverse).unwrap().document,
    ).toStrictEqual(document);
  });

  test("rejecting the reference's deletion gives the note its text back", () => {
    const document = documentOf(false);
    expect(
      resolve(followed(document).document, REVISION_DECISIONS.REJECT),
    ).toStrictEqual(document);
  });

  test("accepting the reference's deletion removes the note", () => {
    expect(
      resolve(followed(documentOf(false)).document, REVISION_DECISIONS.ACCEPT)
        .package.footnotes,
    ).toEqual([]);
  });

  test("refuses without the ids the note's records need", () => {
    const { noteOps } = deleteWithNote(documentOf(false), STAMP, []);
    expect(noteOps.isErr() && noteOps.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
  });

  test("passes over ids the document already holds", () => {
    const { noteOps } = deleteWithNote(documentOf(false), STAMP, [
      STAMP.id,
      ...NEW_IDS,
    ]);
    const edit = applyDocumentOps(documentOf(false), [
      deleteReference(STAMP),
      ...noteOps.unwrap().ops,
    ]).unwrap();
    expect(contractViolation(edit.document)).toBeUndefined();
    expect(noteOps.unwrap().revisions).not.toContain(STAMP.id);
  });

  test("leaves a note another live reference still cites", () => {
    const { noteOps } = deleteWithNote(documentOf(true), STAMP);
    expect(noteOps.unwrap().ops).toEqual([]);
  });

  test("follows the last live reference past another author's earlier deletion", () => {
    const first = deleteReference({ ...STAMP, id: 19, author: "Earlier" });
    const before = applyDocumentOps(documentOf(true), [first]).unwrap()
      .document;
    const at = { story: OP_STORIES.MAIN, blockId: formatParaId(2) } as const;
    const after = applyDocumentOps(before, [
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: { ...at, offset: "Again".length },
        to: { ...at, offset: "Again".length + 1 },
        revision: STAMP,
      },
    ]).unwrap().document;
    const result = noteDeletionOps({
      before,
      after,
      stamp: STAMP,
      revisionIds: NEW_IDS,
    }).unwrap();
    expect(result.ops).toHaveLength(2);
    const edit = applyDocumentOps(after, result.ops).unwrap();
    expect(noteDeletions(edit.document)).toMatchObject({
      live: false,
      deletions: [{ author: STAMP.author }, { author: STAMP.author }],
    });
  });

  test("refuses to track the note removal when the note holds a point comment", () => {
    const document = documentOf(false);
    document.package.document.comments = [
      { id: 12, author: "Reviewer", anchorKind: "point", content: [] },
    ];
    noteParagraphs(document)
      .at(0)
      ?.content.push({ type: "commentReference", id: 12 });
    const { noteOps } = deleteWithNote(document, STAMP);
    expect(noteOps.isErr() && noteOps.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
  });

  test("leaves the note of a suggestion, which stays its author's until accepted", () => {
    const { noteOps } = deleteWithNote(documentOf(false), {
      ...STAMP,
      suggestionId: "suggestion-1",
    });
    expect(noteOps.unwrap().ops).toEqual([]);
  });
});
