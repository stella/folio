import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { validateOpsDocument } from "../contract";
import { formatParaId } from "../idLease";
import { findStoryParagraph } from "../paragraphLocator";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

test("block replacement uses indexed paragraph locations and keeps its exact inverse", () => {
  const paragraph = (id: number): Paragraph => ({
    type: "paragraph",
    paraId: formatParaId(id).toLowerCase(),
    content: [{ type: "run", content: [{ type: "text", text: String(id) }] }],
  });
  const removed = paragraph(170);
  const retained = paragraph(171);
  const prefix = Array.from({ length: 169 }, (_, index) =>
    paragraph(index + 1),
  );
  const body = { content: [...prefix, removed, retained, paragraph(172)] };
  const document: Document = { package: { document: body } };
  expect(validateOpsDocument(document).isOk()).toBe(true);
  expect(findStoryParagraph(body, "000000AA")?.paragraph).toBe(removed);

  const walks = spyOn(blocks, "storyParagraphs");
  try {
    const edit = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [removed, retained],
      blocks: [retained],
    }).unwrap();
    expect(walks.mock.calls.some(([seen]) => seen === body)).toBe(false);
    expect(edit.touched).toEqual({
      inserted: [],
      modified: ["000000ab"],
      removed: ["000000aa"],
    });
    expect(edit.inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: OP_STORIES.MAIN,
        expected: [retained],
        blocks: [removed, retained],
      },
    ]);
    expect(
      applyDocumentOps(edit.document, edit.inverse).unwrap().document,
    ).toEqual(document);
  } finally {
    walks.mockRestore();
  }
});
