import { expect, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import {
  type Document,
  finalizeParagraphFormatting,
  type Paragraph,
  type ParagraphFormatting,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import { paragraphLength } from "../offsets";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const paragraph = ({
  id,
  text,
  formatting,
}: {
  id: string;
  text: string;
  formatting?: ParagraphFormatting;
}): Paragraph => {
  // A parsed source and block insertion both represent an empty pPr as absent.
  const properties = finalizeParagraphFormatting(formatting);
  return {
    type: "paragraph",
    paraId: id,
    content: [{ type: "run", content: [{ type: "text", text }] }],
    ...(properties === undefined ? {} : { formatting: properties }),
  };
};

const formats = [
  undefined,
  {},
  { alignment: "right" },
  { numPr: { kind: "reference", numId: 1, ilvl: 0 }, indentLeft: 720 },
  { styleId: "Heading", spaceBefore: 240 },
] as const satisfies readonly (ParagraphFormatting | undefined)[];

for (const container of ["body", "cell"] as const)
  test(`cancelling an inserted mark tracks its property transfer in a ${container}`, () => {
    for (const sourceFormatting of formats)
      for (const targetFormatting of formats)
        for (const route of ["range", "join"] as const) {
          const anchor = paragraph({ id: "10000001", text: "anchor" });
          const tail = paragraph({
            id: "10000002",
            text: "tail",
            formatting: targetFormatting,
          });
          const content = [anchor, tail];
          const original = {
            package: {
              document: {
                content:
                  container === "body"
                    ? content
                    : [
                        {
                          type: "table",
                          rows: [
                            {
                              type: "tableRow",
                              cells: [{ type: "tableCell", content }],
                            },
                          ],
                        },
                      ],
              },
              styles: {
                styles: [
                  {
                    type: "paragraph",
                    styleId: "Heading",
                    pPr: { alignment: "center" },
                  },
                ],
              },
              numbering: {
                abstractNums: [
                  {
                    abstractNumId: 1,
                    levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
                  },
                ],
                nums: [{ numId: 1, abstractNumId: 1 }],
              },
            },
          } satisfies Document;
          const added = paragraph({
            id: "10000003",
            text: "added",
            formatting: sourceFormatting,
          });
          const insert = {
            type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
            story: OP_STORIES.MAIN,
            at: { type: "before", blockId: tail.paraId ?? "" },
            blocks: [added],
          } as const;
          const inserted = applyDocumentOp(original, {
            ...insert,
            revision: {
              id: 10,
              author: "Reviewer",
              date: "2026-01-01T00:00:00Z",
            },
            newIds: {
              revision: Array.from({ length: 20 }, (_, index) => 11 + index),
            },
          });
          if (inserted.isErr()) throw inserted.error;
          const intent =
            route === "range"
              ? ({
                  type: "replaceText",
                  from: {
                    story: OP_STORIES.MAIN,
                    blockId: added.paraId ?? "",
                    offset: paragraphLength(added),
                  },
                  to: {
                    story: OP_STORIES.MAIN,
                    blockId: tail.paraId ?? "",
                    offset: paragraphLength(tail),
                  },
                  text: "",
                } as const satisfies EditorIntent)
              : ({
                  type: "joinParagraphs",
                  story: OP_STORIES.MAIN,
                  blockId: added.paraId ?? "",
                  nextBlockId: tail.paraId ?? "",
                } as const satisfies EditorIntent);
          const compiled = compileEditorIntent(inserted.value.document, {
            intent,
            mode: {
              type: "suggesting",
              revision: {
                id: 100,
                author: "Reviewer",
                date: "2026-01-02T00:00:00Z",
              },
              newIds: {
                revision: Array.from({ length: 32 }, (_, index) => 101 + index),
              },
            },
          });
          if (compiled.isErr()) throw compiled.error;
          const changed = applyDocumentOps(
            inserted.value.document,
            JSON.parse(JSON.stringify(compiled.value.ops)),
          );
          if (changed.isErr()) throw changed.error;
          const undone = applyDocumentOps(
            changed.value.document,
            JSON.parse(JSON.stringify(changed.value.inverse)),
          );
          if (undone.isErr()) throw undone.error;
          assertExactModel(undone.value.document, inserted.value.document);
          const revisions = [
            ...new Set([
              ...inserted.value.revisions,
              ...changed.value.revisions,
            ]),
          ];
          expect(revisions.length).toBeGreaterThan(0);
          const resolve = (decision: "accept" | "reject") => {
            const result = applyDocumentOp(changed.value.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: revisions,
              decision,
            });
            if (result.isErr()) throw result.error;
            return result.value.document;
          };
          assertExactModel(
            resolve(REVISION_DECISIONS.REJECT).package.document.content,
            original.package.document.content,
          );
          const directInsertion = applyDocumentOp(original, insert);
          if (directInsertion.isErr()) throw directInsertion.error;
          const directPlan = compileEditorIntent(
            directInsertion.value.document,
            { intent, mode: { type: "editing" } },
          );
          if (directPlan.isErr()) throw directPlan.error;
          const direct = applyDocumentOps(
            directInsertion.value.document,
            directPlan.value.ops,
          );
          if (direct.isErr()) throw direct.error;
          assertExactModel(
            resolve(REVISION_DECISIONS.ACCEPT).package.document.content,
            direct.value.document.package.document.content,
          );
        }
  });
