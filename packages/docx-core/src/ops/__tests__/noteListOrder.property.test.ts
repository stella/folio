import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty } from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const paragraph = (id: number): Paragraph => ({
  type: "paragraph",
  paraId: id.toString(16).padStart(8, "0"),
  content: [],
});

test("independent note insertions commute over authored lists and undo exactly", () => {
  assertProperty(
    fc.property(
      fc.uniqueArray(fc.integer({ min: 1, max: 30 }), {
        minLength: 2,
        maxLength: 8,
      }),
      ([firstId, secondId, ...existingIds]) => {
        if (firstId === undefined || secondId === undefined)
          throw new Error("Two fresh note ids are required.");
        const document = {
          package: {
            document: { content: [paragraph(1), paragraph(2)] },
            footnotes: existingIds.map((id) => ({
              type: "footnote",
              id,
              content: [paragraph(100 + id)],
            })),
            endnotes: existingIds.map((id) => ({
              type: "endnote",
              id,
              content: [paragraph(200 + id)],
            })),
          },
        } satisfies Document;
        for (const kind of ["footnote", "endnote"] as const) {
          const add = (id: number, blockId: string) =>
            ({
              type: DOCUMENT_OP_TYPES.ADD_NOTE,
              at: { story: OP_STORIES.MAIN, blockId, offset: 0 },
              note: { type: kind, id, content: [paragraph(300 + id)] },
            }) satisfies DocumentOp;
          const first = add(firstId, "00000001");
          const second = add(secondId, "00000002");
          const forward = applyDocumentOps(document, [first, second]).unwrap();
          const reverse = applyDocumentOps(document, [second, first]).unwrap();
          expect(forward.document).toStrictEqual(reverse.document);
          const notes =
            forward.document.package[
              kind === "footnote" ? "footnotes" : "endnotes"
            ];
          expect(
            notes
              ?.map(({ id }) => id)
              .filter((id) => id !== firstId && id !== secondId),
          ).toEqual(existingIds);
          const undone = applyDocumentOps(
            forward.document,
            JSON.parse(JSON.stringify(forward.inverse)),
          ).unwrap();
          expect(undone.document).toStrictEqual(document);
          const redone = applyDocumentOps(
            undone.document,
            JSON.parse(JSON.stringify(undone.inverse)),
          ).unwrap();
          expect(redone.document).toStrictEqual(forward.document);
        }
      },
    ),
    {
      seed: 20261116,
      numRuns: 80,
      examples: [[[8, 3, 9, 1]]],
      id: "independent note insertions commute and undo over authored lists",
    },
  );
});
