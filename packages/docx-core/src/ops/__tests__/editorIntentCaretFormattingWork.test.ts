import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { validateOpsDocument } from "../contract";
import { compileEditorIntent } from "../editorIntent";
import { formatParaId } from "../idLease";
import * as leaves from "../leaves";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const paragraph = (id: number, text: string, bold = false): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content:
    text === ""
      ? []
      : [
          {
            type: "run",
            formatting: { bold },
            content: [{ type: "text", text }],
          },
        ],
});

test.each(["del", "moveFrom", "ins"] as const)(
  "a caret reads only its joined predecessors for a %s paragraph mark",
  (kind) => {
    const unrelated = Array.from({ length: 128 }, (_, index) =>
      paragraph(index + 10, "Elsewhere"),
    );
    const info = { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
    const left: Paragraph = {
      ...paragraph(1, "Left", true),
      pPrMark: { kind, info },
    };
    const empty: Paragraph = {
      ...paragraph(2, ""),
      pPrMark: { kind: "del", info: { ...info, id: 2 } },
    };
    const target = paragraph(3, "Right");
    const document: Document = {
      package: { document: { content: [...unrelated, left, empty, target] } },
    };
    const from = {
      story: OP_STORIES.MAIN,
      blockId: formatParaId(3),
      offset: 0,
    };
    validateOpsDocument(document).unwrap();
    const spans = spyOn(leaves, "leafSpans");
    try {
      const edit = compileEditorIntent(document, {
        intent: { type: "insertAtom", from, to: from, atom: { type: "tab" } },
        mode: { type: "editing" },
      }).unwrap();
      const op = edit.ops.find(
        (item) => item.type === DOCUMENT_OP_TYPES.INSERT_CONTENT,
      );
      expect(
        op?.type === DOCUMENT_OP_TYPES.INSERT_CONTENT && op.slice.content.at(0),
      ).toMatchObject({
        type: "run",
        formatting: { bold: kind !== "ins" },
      });
      expect(
        spans.mock.calls.filter(([content]) =>
          unrelated.some((item) => item.content === content),
        ),
      ).toHaveLength(0);
    } finally {
      spans.mockRestore();
    }
  },
);
