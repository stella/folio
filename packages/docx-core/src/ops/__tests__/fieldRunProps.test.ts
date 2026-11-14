import { describe, expect, test } from "bun:test";

import type {
  ComplexField,
  Document,
  Paragraph,
  Run,
  SimpleField,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { ownRunPropertyChanges } from "../leaves";
import { findStoryBody } from "../stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
  REVISION_DECISIONS,
  type SetRunPropsOp,
} from "../types";

const run = (text: string, formatting?: Run["formatting"]): Run => ({
  type: "run",
  ...(formatting === undefined ? {} : { formatting }),
  content: [{ type: "text", text }],
});

const simpleField = (result: readonly Run[]): SimpleField => ({
  type: "simpleField",
  instruction: "PAGE",
  fieldType: "PAGE",
  content: [...result],
});

const complexField = (result: readonly Run[]): ComplexField => ({
  type: "complexField",
  instruction: " NUMPAGES ",
  fieldType: "NUMPAGES",
  fieldCode: [
    {
      type: "run",
      formatting: { italic: true },
      content: [{ type: "instrText", text: " NUMPAGES " }],
    },
  ],
  fieldResult: [...result],
  formatting: { italic: true },
});

const documentOf = (content: Paragraph["content"]): Document => ({
  package: {
    document: { content: [{ type: "paragraph", paraId: "00000001", content }] },
  },
});

const at = (offset: number, zeroWidthBefore?: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset,
  ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
});

const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const undone = (document: Document, inverse: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, inverse);
  if (result.isErr()) throw result.error;
  return result.value.document;
};

const contentOf = (document: Document, story: OpStory = OP_STORIES.MAIN) => {
  const paragraph = findStoryBody(document, story)?.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("No paragraph");
  return paragraph.content;
};

const fieldsOf = (document: Document) =>
  contentOf(document).flatMap((item) =>
    item.type === "simpleField" || item.type === "complexField" ? [item] : [],
  );

const resultFormatting = (document: Document) =>
  fieldsOf(document).map((field) =>
    (field.type === "simpleField" ? field.content : field.fieldResult).map(
      (item) => (item.type === "run" ? item.formatting : undefined),
    ),
  );

const revision = { id: 71, author: "Reviewer", date: "2026-11-14T00:00:00Z" };

// Each run a tracked patch changes records its own revision id: the first
// keeps the stamp's, the others take these.
const newIds = { revision: [201, 202, 203, 204, 205, 206] };

/** The run property change ids of a paragraph's runs and field results, in document order. */
const changeIds = (document: Document, story: OpStory = OP_STORIES.MAIN) =>
  contentOf(document, story).flatMap((item) =>
    ownRunPropertyChanges(item).map(({ info }) => info.id),
  );

const bold = (from: number, to: number): SetRunPropsOp => ({
  type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
  from: at(from),
  to: at(to),
  patch: { bold: true },
});

// "ab" | PAGE "7" | "cd" | NUMPAGES "1" "2" | "ef": offsets 0-2, 2-3, 3-5, 5-6, 6-8.
const original = documentOf([
  run("ab"),
  simpleField([run("7", { italic: true })]),
  run("cd"),
  complexField([run("1", { italic: true }), run("2", { italic: true })]),
  run("ef"),
]);

describe("run properties over a field", () => {
  test("a patch over a field's unit reaches the runs of its result and not its instruction", () => {
    const { document } = applied(original, bold(0, 8));
    expect(resultFormatting(document)).toEqual([
      [{ italic: true, bold: true }],
      [
        { italic: true, bold: true },
        { italic: true, bold: true },
      ],
    ]);
    const complex = fieldsOf(document).at(1);
    expect(complex?.type === "complexField" && complex.fieldCode).toEqual(
      complexField([]).fieldCode,
    );
  });

  test("a range ending at a field's end reaches it; one ending at its start or starting at its end does not", () => {
    expect(resultFormatting(applied(original, bold(1, 3)).document)).toEqual([
      [{ italic: true, bold: true }],
      [{ italic: true }, { italic: true }],
    ]);
    expect(resultFormatting(applied(original, bold(0, 2)).document)).toEqual(
      resultFormatting(original),
    );
    expect(resultFormatting(applied(original, bold(3, 5)).document)).toEqual(
      resultFormatting(original),
    );
  });

  test("the inverse restores the paragraph exactly, also when the result runs differed", () => {
    const mixed = documentOf([
      run("ab"),
      complexField([run("1", { bold: true }), run("2", { italic: true })]),
    ]);
    for (const [document, op] of [
      [original, bold(0, 8)],
      [original, bold(2, 3)],
      [mixed, bold(0, 3)],
    ] as const) {
      const { document: next, inverse } = applied(document, op);
      expect(next).not.toEqual(document);
      expect(undone(next, inverse)).toStrictEqual(document);
    }
  });

  test("a tracked patch records the change on each result run; reject restores and accept keeps it", () => {
    const op = { ...bold(2, 6), revision, newIds };
    const direct = applied(original, bold(2, 6)).document;
    const { document: tracked, inverse } = applied(original, op);
    expect(
      fieldsOf(tracked).flatMap((field) =>
        (field.type === "simpleField" ? field.content : field.fieldResult).map(
          (item) =>
            item.type === "run"
              ? item.propertyChanges?.map(
                  ({ previousFormatting }) => previousFormatting,
                )
              : undefined,
        ),
      ),
    ).toEqual([[{ italic: true }], [{ italic: true }], [{ italic: true }]]);
    // The field result runs and the run between them: four records, four ids.
    const ids = changeIds(tracked);
    expect(ids).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
    expect(ids).toContain(revision.id);
    expect(undone(tracked, inverse)).toStrictEqual(original);
    const resolved = (decision: "accept" | "reject") =>
      applied(tracked, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: ids,
        decision:
          decision === "accept"
            ? REVISION_DECISIONS.ACCEPT
            : REVISION_DECISIONS.REJECT,
      }).document;
    expect(resolved("reject")).toStrictEqual(original);
    expect(resolved("accept")).toStrictEqual(direct);
  });

  test("a field without a result takes a direct patch on its own formatting and keeps it under a tracked one", () => {
    const resultless = documentOf([run("ab"), complexField([])]);
    const { document, inverse } = applied(resultless, bold(0, 3));
    const field = fieldsOf(document).at(0);
    expect(field?.type === "complexField" && field.formatting).toEqual({
      italic: true,
      bold: true,
    });
    expect(undone(document, inverse)).toStrictEqual(resultless);
    const tracked = applied(resultless, { ...bold(0, 3), revision }).document;
    expect(fieldsOf(tracked)).toEqual(fieldsOf(resultless));
  });
});

describe("resolving tracked formats inside field results", () => {
  test("one batch, two batches and the other order agree, accepting and rejecting", () => {
    const second = { ...revision, id: 72 };
    let tracked = applied(original, {
      ...bold(0, 3),
      revision,
      newIds,
    }).document;
    const first = changeIds(tracked);
    tracked = applied(tracked, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at(5),
      to: at(8),
      patch: { strike: true },
      revision: second,
      newIds: { revision: [301, 302, 303] },
    }).document;
    const later = changeIds(tracked).filter((id) => !first.includes(id));
    expect(first).toHaveLength(2);
    expect(later).toHaveLength(3);
    const resolve = (
      document: Document,
      ids: number[],
      decision: "accept" | "reject",
    ) =>
      applied(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: ids,
        decision:
          decision === "accept"
            ? REVISION_DECISIONS.ACCEPT
            : REVISION_DECISIONS.REJECT,
      }).document;
    for (const decision of ["accept", "reject"] as const) {
      const batch = resolve(tracked, [...first, ...later], decision);
      const inTurn = resolve(
        resolve(tracked, first, decision),
        later,
        decision,
      );
      const otherOrder = resolve(
        resolve(tracked, later, decision),
        first,
        decision,
      );
      expect(inTurn).toStrictEqual(batch);
      expect(otherOrder).toStrictEqual(batch);
      expect(
        fieldsOf(batch).every((field) =>
          (field.type === "simpleField"
            ? field.content
            : field.fieldResult
          ).every(
            (item) => item.type !== "run" || item.propertyChanges === undefined,
          ),
        ),
      ).toBe(true);
      if (decision === "reject") expect(batch).toStrictEqual(original);
    }
  });
});

describe("run properties over a note's own reference mark", () => {
  const story = { kind: "footnote", id: 1 } as const satisfies OpStory;
  const marker: Run = {
    type: "run",
    formatting: { styleId: "FootnoteReference" },
    content: [{ type: "noteMarker", kind: "footnote" }],
  };
  const withNote = applied(documentOf([run("abcd")]), {
    type: DOCUMENT_OP_TYPES.ADD_NOTE,
    at: at(2),
    note: {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          paraId: "00000002",
          content: [marker, run("note")],
        },
      ],
    },
  }).document;
  const position = (offset: number, zeroWidthBefore?: number) => ({
    story,
    blockId: "00000002",
    offset,
    ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
  });
  const markerFormatting = (document: Document) =>
    contentOf(document, story).flatMap((item) =>
      item.type === "run" &&
      item.content.some((content) => content.type === "noteMarker")
        ? [{ formatting: item.formatting, changes: item.propertyChanges }]
        : [],
    );

  test("a range holding the mark reaches its run, direct and tracked, and both undo", () => {
    for (const tracked of [false, true]) {
      const op: SetRunPropsOp = {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: position(0, 0),
        to: position(4),
        patch: { bold: true },
        ...(tracked ? { revision, newIds } : {}),
      };
      const { document, inverse } = applied(withNote, op);
      expect(markerFormatting(document)).toEqual([
        {
          formatting: { styleId: "FootnoteReference", bold: true },
          changes: tracked
            ? [
                {
                  type: "runPropertyChange",
                  info: revision,
                  boundaryJoins: [],
                  previousFormatting: { styleId: "FootnoteReference" },
                },
              ]
            : undefined,
        },
      ]);
      expect(undone(document, inverse)).toStrictEqual(withNote);
    }
  });

  test("a range starting after the mark leaves its run", () => {
    const { document } = applied(withNote, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: position(0),
      to: position(4),
      patch: { bold: true },
    });
    expect(markerFormatting(document)).toEqual(markerFormatting(withNote));
  });
});
