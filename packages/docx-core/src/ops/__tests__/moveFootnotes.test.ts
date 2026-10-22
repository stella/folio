import { expect, test } from "bun:test";

import type { Document, Paragraph, Run } from "../../model/document";
import { applyDocumentOps } from "../apply";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  paragraphVisibleText,
  type EditorIntent,
} from "../editorIntent";
import { OP_STORIES } from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const reference = (id: number): Run => ({
  type: "run",
  content: [{ type: "footnoteRef", id }],
});

const paragraph = (paraId: string, content: Paragraph["content"]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

test("one move preserves two footnotes referenced from separate paragraphs and undoes exactly", () => {
  const firstReference = reference(1);
  const secondReference = reference(2);
  const document = {
    package: {
      document: {
        content: [
          paragraph("00000001", [run("A"), firstReference]),
          paragraph("00000002", [secondReference, run("B")]),
          paragraph("00000003", [run("T")]),
        ],
      },
      footnotes: [
        { type: "footnote", id: 1, content: [paragraph("00000011", [run("first note")])] },
        { type: "footnote", id: 2, content: [paragraph("00000012", [run("second note")])] },
      ],
    },
  } satisfies Document;
  const intent = {
    type: "moveFragment",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
    to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 1 },
    target: { story: OP_STORIES.MAIN, blockId: "00000003", offset: 1 },
    paragraphs: [
      { type: "paragraph", content: [firstReference] },
      { type: "paragraph", content: [secondReference] },
    ],
    openStart: 1,
    openEnd: 1,
  } as const satisfies EditorIntent;
  const allocation = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: allocation.newIds },
  }).unwrap();

  const moved = applyDocumentOps(document, compiled.ops).unwrap();
  expect(moved.document.package.footnotes).toStrictEqual(document.package.footnotes);
  expect(
    moved.document.package.document.content.map((block) =>
      block.type === "paragraph" ? paragraphVisibleText(block) : "",
    ),
  ).toEqual(["AB", "T\uFFFC", "\uFFFC"]);
  const targetContent = moved.document.package.document.content
    .slice(1)
    .flatMap((block) => (block.type === "paragraph" ? block.content : []));
  expect(targetContent).toEqual([run("T"), firstReference, secondReference]);
  expect(applyDocumentOps(moved.document, moved.inverse).unwrap().document).toStrictEqual(document);
});
