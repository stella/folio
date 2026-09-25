import { describe, expect, test } from "bun:test";

import type { Document, Insertion, Paragraph, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { replaceParagraphs, storyParagraphs } from "../blocks";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const trackedRun = (text: string, id: number): Run => ({
  type: "run",
  content: [{ type: "text", text }],
  propertyChanges: [
    { type: "runPropertyChange", info: { id, author: "A" }, previousFormatting: {} },
  ],
});

const insertion = (id: number, content: Run[]): Insertion => ({
  type: "insertion",
  info: { id, author: "A", date: "2026-01-02T03:04:05Z" },
  content,
});

const documentOf = (...content: Paragraph[]): Document => ({ package: { document: { content } } });

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const applied = (document: Document, op: DocumentOp): AppliedDocumentOp => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const reasonOf = (document: Document, ops: readonly DocumentOp[]): string | undefined => {
  const result = applyDocumentOps(document, ops);
  return result.isErr() ? result.error.reason : undefined;
};

const paragraphs = (document: Document): Paragraph[] =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph);

describe("an undo that no longer applies", () => {
  test("a formatting undo after a later change to the same property is stale", () => {
    const document = documentOf({ type: "paragraph", paraId: "00000001", content: [run("a")] });
    const center = applied(document, {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      patch: { alignment: "center" },
    });
    const end = applied(center.document, {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      patch: { alignment: "end" },
    });
    expect(reasonOf(end.document, center.inverse)).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });

  test("a run formatting undo after a later change to the same property is stale", () => {
    const document = documentOf({ type: "paragraph", paraId: "00000001", content: [run("abc")] });
    const bold = applied(document, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at("00000001", 0),
      to: at("00000001", 3),
      patch: { bold: true },
    });
    const unbold = applied(bold.document, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at("00000001", 1),
      to: at("00000001", 2),
      patch: { bold: false },
    });
    expect(reasonOf(unbold.document, bold.inverse)).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });

  test("undoing a split after the new paragraph's properties changed is stale", () => {
    const document = documentOf({ type: "paragraph", paraId: "00000001", content: [run("ab")] });
    const split = applied(document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 1),
      newBlockId: "0000000A",
    });
    const restyled = applied(split.document, {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "0000000A",
      patch: { styleId: "Heading1" },
    });
    expect(reasonOf(restyled.document, split.inverse)).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });

  test("renumbering or pagination recomputed since does not make an undo stale", () => {
    const document = documentOf({ type: "paragraph", paraId: "00000001", content: [run("ab")] });
    const split = applied(document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 1),
      newBlockId: "0000000A",
    });
    const [, second] = storyParagraphs(split.document.package.document);
    if (second === undefined) throw new Error("fixture");
    // What a numbering and layout pass writes back.
    const relaidOut = replaceParagraphs({
      document: split.document,
      story: OP_STORIES.MAIN,
      at: second,
      count: 1,
      replacement: [
        {
          ...second.paragraph,
          listRendering: { marker: "2.", level: 0, numId: 1, isBullet: false },
          renderedPageBreakBefore: true,
        },
      ],
    });
    expect(reasonOf(relaidOut, split.inverse)).toBeUndefined();
  });
});

describe("records carrying ids", () => {
  const trackedDocument = documentOf({
    type: "paragraph",
    paraId: "00000001",
    content: [insertion(7, [trackedRun("abcd", 8)])],
  });

  test("a split inside a tracked insertion continues it under the new ids it names", () => {
    const op: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 2),
      newBlockId: "0000000A",
      newIds: { revision: [70, 80] },
    };
    const split = applied(trackedDocument, op);
    expect(paragraphs(split.document).map(({ content }) => content)).toEqual([
      [insertion(7, [trackedRun("ab", 8)])],
      [insertion(70, [trackedRun("cd", 80)])],
    ]);
    expect(reasonOf(split.document, split.inverse)).toBeUndefined();
    const undone = applyDocumentOps(split.document, split.inverse);
    expect(undone.isOk() ? undone.value.document : undefined).toEqual(trackedDocument);
  });

  test("a split inside a tracked insertion without new ids is refused", () => {
    expect(
      reasonOf(trackedDocument, [
        { type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, at: at("00000001", 2), newBlockId: "0000000A" },
      ]),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS);
  });

  test("patching part of a run with a tracked property change gives the cut pieces new ids", () => {
    const document = documentOf({
      type: "paragraph",
      paraId: "00000001",
      content: [trackedRun("abc", 8)],
    });
    const patched = applied(document, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at("00000001", 1),
      to: at("00000001", 2),
      patch: { bold: true },
      newIds: { revision: [81, 82] },
    });
    const ids = paragraphs(patched.document)[0]?.content.map((item) =>
      item.type === "run" ? item.propertyChanges?.[0]?.info.id : undefined,
    );
    expect(ids).toEqual([8, 81, 82]);
    const undone = applyDocumentOps(patched.document, patched.inverse);
    expect(undone.isOk() ? undone.value.document : undefined).toEqual(document);
  });
});
