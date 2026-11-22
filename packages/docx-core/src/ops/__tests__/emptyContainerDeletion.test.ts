/**
 * A tracked deletion marks runs. An empty hyperlink, content control or
 * wrapper holds none: a deletion of nothing else would be a record no file
 * or reader shows, so it is refused, and a suggestion plans none for it.
 */

import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent } from "../editorIntent";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionStamp,
} from "../types";

const STAMP: RevisionStamp = {
  id: 2,
  author: "Reviewer",
  date: "2026-01-02T03:04:05Z",
};

const EMPTY_CONTAINERS: Record<string, ParagraphContent> = {
  hyperlink: {
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children: [],
  },
  "content control": {
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 5, tag: "clause" },
    content: [],
  },
  "bidirectional wrapper": {
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content: [],
  },
};

const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
        },
        { type: "paragraph", paraId: "00000002", content },
      ],
    },
  },
});

const gap = (offset: number, zeroWidthBefore: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000002",
  offset,
  zeroWidthBefore,
});

describe.each(Object.entries(EMPTY_CONTAINERS))(
  "an empty %s",
  (_, container) => {
    test("a tracked deletion of it alone is refused; a direct one removes it", () => {
      const document = documentOf([container]);
      const op = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: gap(0, 0),
        to: gap(0, 1),
      } as const;
      const tracked = applyDocumentOp(document, { ...op, revision: STAMP });
      expect(tracked.isErr() && tracked.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
      const direct = applyDocumentOp(document, op);
      if (direct.isErr()) throw direct.error;
      expect(direct.value.document.package.document.content[1]).toMatchObject({
        content: [],
      });
    });

    test("a tracked deletion of it beside text records the text", () => {
      const document = documentOf([
        container,
        { type: "run", content: [{ type: "text", text: "b" }] },
      ]);
      const op = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: gap(0, 0),
        to: gap(1, 0),
      } as const;
      const tracked = applyDocumentOp(document, { ...op, revision: STAMP });
      if (tracked.isErr()) throw tracked.error;
      expect(tracked.value.revisions).toEqual([STAMP.id]);
      const accepted = applyDocumentOp(tracked.value.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: tracked.value.revisions,
        decision: REVISION_DECISIONS.ACCEPT,
      });
      const direct = applyDocumentOp(document, op);
      if (accepted.isErr()) throw accepted.error;
      if (direct.isErr()) throw direct.error;
      expect(accepted.value.document).toStrictEqual(direct.value.document);
    });

    test("a suggested removal of its paragraph records only the join", () => {
      const document = documentOf([container]);
      const compiled = compileEditorIntent(document, {
        intent: {
          type: "replaceText",
          from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
          to: gap(0, 1),
          text: "",
        },
        mode: {
          type: "suggesting",
          revision: STAMP,
          newIds: { revision: [3, 4], control: [] },
        },
      });
      if (compiled.isErr()) throw compiled.error;
      expect(
        compiled.value.ops.some(
          (op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE,
        ),
      ).toBe(false);
      const applied = applyDocumentOps(document, compiled.value.ops);
      if (applied.isErr()) throw applied.error;
      const [first] = applied.value.document.package.document.content;
      expect(first).toMatchObject({ pPrMark: { kind: "del" } });
    });

    test.each([
      ["standing in the paragraph", (node: ParagraphContent) => node],
      [
        "inside the suggesting author's own insertion",
        (node: ParagraphContent): ParagraphContent => ({
          type: "insertion",
          info: { id: 7, author: STAMP.author, date: STAMP.date },
          content: [node],
        }),
      ],
    ])(
      "a replacement of text around it, apart from it, keeps it %s in both modes",
      (_where, place) => {
        // The bookmark ends the stretch the text is in: the container stands
        // alone in its own, which a suggestion marks nothing in.
        const document = documentOf([
          { type: "run", content: [{ type: "text", text: "a" }] },
          { type: "bookmarkStart", id: 1, name: "_Mark" },
          place(container),
          { type: "bookmarkEnd", id: 1 },
          { type: "run", content: [{ type: "text", text: "b" }] },
        ]);
        const intent = {
          type: "replaceText",
          from: gap(0, 0),
          to: gap(2, 0),
          text: "edited",
        } as const;
        const direct = compileEditorIntent(document, {
          intent,
          mode: { type: "editing" },
        });
        const tracked = compileEditorIntent(document, {
          intent,
          mode: {
            type: "suggesting",
            revision: STAMP,
            newIds: { revision: [3, 4], control: [] },
          },
        });
        if (direct.isErr()) throw direct.error;
        if (tracked.isErr()) throw tracked.error;
        const edited = applyDocumentOps(document, direct.value.ops);
        const suggested = applyDocumentOps(document, tracked.value.ops);
        if (edited.isErr()) throw edited.error;
        if (suggested.isErr()) throw suggested.error;
        const accepted = applyDocumentOp(suggested.value.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: suggested.value.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        });
        if (accepted.isErr()) throw accepted.error;
        expect(accepted.value.document).toStrictEqual(edited.value.document);
        expect(edited.value.document.package.document.content[1]).toMatchObject(
          {
            content: expect.arrayContaining([place(container)]),
          },
        );
      },
    );

    // Deleted content parts the records a deletion is written in, so a
    // container past it shares no record with the text.
    test.each([
      [
        "past a pending deletion",
        (): ParagraphContent[] => [
          { type: "run", content: [{ type: "text", text: "a" }] },
          {
            type: "deletion",
            info: { id: 7, author: "Other", date: STAMP.date },
            content: [{ type: "run", content: [{ type: "text", text: "d" }] }],
          },
        ],
      ],
      [
        "past a pending move",
        (): ParagraphContent[] => [
          { type: "run", content: [{ type: "text", text: "a" }] },
          {
            type: "moveFrom",
            info: { id: 7, author: "Other", date: STAMP.date },
            content: [{ type: "run", content: [{ type: "text", text: "d" }] }],
          },
        ],
      ],
      [
        "past another author's insertion and a pending deletion",
        (): ParagraphContent[] => [
          {
            type: "insertion",
            info: { id: 8, author: "Other", date: STAMP.date },
            content: [{ type: "run", content: [{ type: "text", text: "i" }] }],
          },
          {
            type: "deletion",
            info: { id: 7, author: "Third", date: STAMP.date },
            content: [{ type: "run", content: [{ type: "text", text: "d" }] }],
          },
        ],
      ],
      [
        "past the suggesting author's own insertion and a pending deletion",
        (): ParagraphContent[] => [
          {
            type: "insertion",
            info: { id: 8, author: STAMP.author, date: STAMP.date },
            content: [{ type: "run", content: [{ type: "text", text: "i" }] }],
          },
          {
            type: "deletion",
            info: { id: 7, author: "Other", date: STAMP.date },
            content: [{ type: "run", content: [{ type: "text", text: "d" }] }],
          },
        ],
      ],
    ])(
      "a replacement of text around it keeps it in both modes %s",
      (_where, before) => {
        const document = documentOf([
          ...before(),
          container,
          { type: "bookmarkStart", id: 1, name: "_Mark" },
          { type: "run", content: [{ type: "text", text: "b" }] },
          { type: "bookmarkEnd", id: 1 },
        ]);
        const intent = {
          type: "replaceText",
          from: gap(0, 0),
          to: gap(3, 0),
          text: "edited",
        } as const;
        const direct = compileEditorIntent(document, {
          intent,
          mode: { type: "editing" },
        });
        const tracked = compileEditorIntent(document, {
          intent,
          mode: {
            type: "suggesting",
            revision: STAMP,
            newIds: { revision: [3, 4, 5, 6], control: [] },
          },
        });
        if (direct.isErr()) throw direct.error;
        if (tracked.isErr()) throw tracked.error;
        const edited = applyDocumentOps(document, direct.value.ops);
        const suggested = applyDocumentOps(document, tracked.value.ops);
        if (edited.isErr()) throw edited.error;
        if (suggested.isErr()) throw suggested.error;
        const accepted = applyDocumentOp(suggested.value.document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: suggested.value.revisions,
          decision: REVISION_DECISIONS.ACCEPT,
        });
        if (accepted.isErr()) throw accepted.error;
        expect(accepted.value.document).toStrictEqual(edited.value.document);
        const [, suggestedParagraph] =
          suggested.value.document.package.document.content;
        if (suggestedParagraph?.type !== "paragraph")
          throw new Error("Expected a paragraph");
        expect(suggestedParagraph.content).toContainEqual(container);
        const undone = applyDocumentOps(
          suggested.value.document,
          suggested.value.inverse,
        );
        if (undone.isErr()) throw undone.error;
        expect(undone.value.document).toStrictEqual(document);
      },
    );

    test("a direct deletion of nothing but it still removes it", () => {
      const document = documentOf([container]);
      const compiled = compileEditorIntent(document, {
        intent: {
          type: "replaceText",
          from: gap(0, 0),
          to: gap(0, 1),
          text: "",
        },
        mode: { type: "editing" },
      });
      if (compiled.isErr()) throw compiled.error;
      const applied = applyDocumentOps(document, compiled.value.ops);
      if (applied.isErr()) throw applied.error;
      expect(applied.value.document.package.document.content[1]).toMatchObject({
        content: [],
      });
    });
  },
);
