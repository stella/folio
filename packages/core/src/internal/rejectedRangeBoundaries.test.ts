import { describe, expect, test } from "bun:test";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import { createHarnessState } from "../__tests__/editorHarness";
import { rejectAllChanges } from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

// Rejecting an insertion keeps the range boundaries it held where the
// insertion was, in the editor as in the engine.

const insertion = { id: 3, author: "A", date: "2026-03-04T05:00:03Z" };

const documentWith = (content: Paragraph["content"]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { type: "paragraph", paraId: "00000001", content },
  ];
  return document;
};

/** The paragraph as its text and the range boundaries between it, in order. */
const outline = (content: readonly unknown[]): string[] =>
  JSON.stringify(content)
    .match(/"type":"(?:bookmarkStart|bookmarkEnd)"|"text":"[^"]*"/gu)
    ?.map((token) => token.replace(/^"type":|^"text":/u, "")) ?? [];

const firstParagraphContent = (document: Document) => {
  const first = document.package.document.content[0];
  return first?.type === "paragraph" ? first.content : [];
};

const rejectInEditor = (document: Document) => {
  const state = createHarnessState(document, "editing");
  let resolved = state;
  expect(
    rejectAllChanges()(state, (tr) => {
      resolved = state.apply(tr);
    }),
  ).toBe(true);
  return firstParagraphContent(
    fromProseDoc(resolved.doc, document, {
      stylesheetSource: { type: "package" },
    }),
  );
};

const rejectInEngine = (document: Document) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [insertion.id],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return firstParagraphContent(applied.value.document);
};

describe("rejecting an insertion keeps its range boundaries", () => {
  test("a bookmark inside inserted text stays where the insertion was, in both", () => {
    const document = documentWith([
      { type: "run", content: [{ type: "text", text: "a" }] },
      {
        type: "insertion",
        info: insertion,
        content: [
          { type: "run", content: [{ type: "text", text: "x" }] },
          { type: "bookmarkStart", id: 5, name: "_Ref5" },
          { type: "bookmarkEnd", id: 5 },
          { type: "run", content: [{ type: "text", text: "y" }] },
        ],
      },
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
    const editor = outline(rejectInEditor(document));
    expect(editor).toEqual(outline(rejectInEngine(document)));
    expect(editor).toEqual(['"a"', '"bookmarkStart"', '"bookmarkEnd"', '"b"']);
  });
});
