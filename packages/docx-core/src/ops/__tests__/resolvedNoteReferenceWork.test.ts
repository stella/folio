import { expect, spyOn, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { formatParaId } from "../idLease";
import * as noteMarks from "../noteMarks";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const paragraph = (id: number, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content,
});
const text: ParagraphContent = {
  type: "run",
  content: [{ type: "text", text: "stay" }],
};
const noteRef = (type: "footnoteRef" | "endnoteRef", id: number): Run => ({
  type: "run",
  content: [{ type, id }],
});

for (const decision of Object.values(REVISION_DECISIONS)) {
  test.each([4, 16])(
    `one ${decision} decision checks surviving references together for %i note pairs`,
    (count) => {
      const ids = Array.from({ length: count }, (_, index) => index + 1);
      const refs = ids.flatMap((id) => [
        noteRef("footnoteRef", id),
        noteRef("endnoteRef", id),
      ]);
      const untouched = paragraph(10, [text]);
      const live = ids.map((id) =>
        noteRef(id % 2 === 0 ? "footnoteRef" : "endnoteRef", id),
      );
      const document: Document = {
        package: {
          document: {
            content: [
              untouched,
              paragraph(11, [
                {
                  type:
                    decision === REVISION_DECISIONS.ACCEPT
                      ? "deletion"
                      : "insertion",
                  info: { id: 500, author: "Author" },
                  content: refs,
                },
              ]),
              paragraph(12, live),
            ],
          },
          footnotes: ids.map((id) => ({
            type: "footnote",
            id,
            content: [paragraph(100 + id, [text])],
          })),
          endnotes: ids.map((id) => ({
            type: "endnote",
            id,
            content: [paragraph(200 + id, [text])],
          })),
        },
      };
      const scans = spyOn(noteMarks, "noteReferenceSpans");
      try {
        const edit = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [500],
          decision,
        }).unwrap();
        expect(edit.document.package.footnotes?.map(({ id }) => id)).toEqual(
          ids.filter((id) => id % 2 === 0),
        );
        expect(edit.document.package.endnotes?.map(({ id }) => id)).toEqual(
          ids.filter((id) => id % 2 === 1),
        );
        // One removed-reference scan and one shared note census visit this paragraph.
        expect(
          scans.mock.calls.filter(([content]) => content === untouched.content),
        ).toHaveLength(2);
        expect(
          applyDocumentOps(edit.document, edit.inverse).unwrap().document,
        ).toEqual(document);
      } finally {
        scans.mockRestore();
      }
    },
  );
}
