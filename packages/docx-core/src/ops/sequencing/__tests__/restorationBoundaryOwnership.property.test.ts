import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, ParagraphContent } from "../../../model/document";
import { applyDocumentOp } from "../../apply";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "../../types";
import { applyBatch } from "../applyBatch";
import type { DocumentBatch, SequencedBatch } from "../envelope";
import { transformBatch } from "../transform";

const at = {
  story: OP_STORIES.MAIN,
  blockId: "10000001",
  offset: 1,
  zeroWidthBefore: 0,
} as const;
const run = (text: string) =>
  ({
    type: "run",
    content: [{ type: "text", text }],
  }) satisfies ParagraphContent;
const batch = (opId: string, ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops,
});
type InOrderOptions = {
  document: Document;
  first: DocumentBatch;
  second: DocumentBatch;
};
const inOrder = ({ document, first, second }: InOrderOptions) => {
  const applied = applyBatch(document, first.ops).unwrap();
  const sequenced = {
    ...first,
    revision: 1,
    effects: applied.effects,
  } satisfies SequencedBatch;
  const moved = transformBatch(second, [
    JSON.parse(JSON.stringify(sequenced)),
  ]).unwrap();
  const result = applyBatch(
    applied.document,
    JSON.parse(JSON.stringify(moved.ops)),
  ).unwrap();
  const undone = applyBatch(
    result.document,
    JSON.parse(JSON.stringify(result.inverse)),
  ).unwrap();
  expect(undone.document).toStrictEqual(applied.document);
  expect(
    applyBatch(undone.document, undone.inverse).unwrap().document,
  ).toStrictEqual(result.document);
  return result.document;
};

test(
  "restoring either end of a revision keeps concurrent edits outside its identity",
  () => {
    assertProperty(
      fc.property(
        fc.array(fc.constantFrom("a", "β", "中"), {
          minLength: 2,
          maxLength: 8,
        }),
        fc.integer({ min: 0, max: 3 }),
        (characters, depth) => {
          let content: ParagraphContent[] = [run("LR")];
          for (let index = 0; index < depth; index += 1)
            content = [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "embedding",
                content,
              },
            ];
          const base = {
            package: {
              document: {
                content: [{ type: "paragraph", paraId: at.blockId, content }],
              },
            },
          } satisfies Document;
          const inserted = applyDocumentOp(base, {
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at,
            slice: {
              openStart: 0,
              openEnd: 0,
              content: [run(characters.join(""))],
            },
            revision: {
              id: 7,
              author: "Reviewer",
              date: "2026-01-01T00:00:00Z",
            },
            newIds: { revision: [], control: [] },
          }).unwrap().document;
          for (const side of ["start", "end"] as const) {
            const offset =
              side === "start" ? at.offset : at.offset + characters.length - 1;
            const deleted = applyDocumentOp(inserted, {
              type: DOCUMENT_OP_TYPES.DELETE_RANGE,
              from: { ...at, offset },
              to: { ...at, offset: offset + 1 },
              newIds: { revision: [], control: [] },
            }).unwrap();
            if (
              deleted.inverse.at(0)?.type !== DOCUMENT_OP_TYPES.REPLACE_INLINE
            )
              panic(
                "A partial revision deletion must produce a paragraph restoration.",
              );
            for (const kind of ["typing", "pointComment"] as const) {
              const operationAt = (positionOffset: number) =>
                kind === "typing"
                  ? ({
                      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
                      at: { ...at, offset: positionOffset },
                      slice: { openStart: 0, openEnd: 0, content: [run("x")] },
                      newIds: { revision: [], control: [] },
                    } as const satisfies DocumentOp)
                  : ({
                      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
                      comment: {
                        id: 9,
                        author: "Reviewer",
                        content: [
                          {
                            type: "paragraph",
                            paraId: "20000001",
                            content: [run("Note")],
                          },
                        ],
                      },
                      anchor: {
                        kind: "point",
                        at: { ...at, offset: positionOffset },
                      },
                      newIds: { revision: [], control: [] },
                    } as const satisfies DocumentOp);
              const restore = batch("undo", deleted.inverse);
              const edit = batch(kind, [operationAt(offset)]);
              const before = structuredClone({
                document: deleted.document,
                restore,
                edit,
              });
              const first = inOrder({
                document: deleted.document,
                first: restore,
                second: edit,
              });
              const second = inOrder({
                document: deleted.document,
                first: edit,
                second: restore,
              });
              expect(first).toStrictEqual(second);
              const expected = applyDocumentOp(
                inserted,
                operationAt(
                  side === "start" ? at.offset : at.offset + characters.length,
                ),
              ).unwrap().document;
              expect(first).toStrictEqual(expected);
              expect({
                document: deleted.document,
                restore,
                edit,
              }).toStrictEqual(before);
            }
          }
        },
      ),
      {
        id: "restoration preserves identified source gap ownership",
        numRuns: 50,
      },
    );
  },
  propertyTestTimeout(10_000),
);
