import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import {
  COMMENT_PART_RELATIONSHIPS,
  type Document,
  type Paragraph,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const sourceDocument = (text: string): Document =>
  applyDocumentOp(
    {
      package: {
        document: {
          content: [paragraph("10000001", text), paragraph("10000002", "stay")],
        },
        footnotes: [
          {
            type: "footnote",
            id: 3,
            content: [paragraph("30000003", "untouched footnote")],
          },
        ],
        endnotes: [
          {
            type: "endnote",
            id: 3,
            content: [paragraph("30000004", "untouched endnote")],
          },
        ],
      },
    },
    {
      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      comment: {
        id: 11,
        author: "Reviewer",
        content: [paragraph("2000000B", "unrelated comment")],
      },
      anchor: {
        kind: "point",
        at: { story: "main", blockId: "10000002", offset: 1 },
      },
    },
  ).unwrap().document;
const revision = { id: 41, author: "Reviewer", date: "2026-11-17T00:00:00Z" };
const textArbitrary = fc
  .array(fc.constantFrom("a", "b", "c", " "), { minLength: 2, maxLength: 10 })
  .map((letters) => letters.join(""));

const expectRestoration = (document: Document, ops: readonly DocumentOp[]) => {
  const original = structuredClone(document);
  const applied = applyDocumentOps(document, ops).unwrap();
  expect(contractViolation(applied.document)).toBeUndefined();
  const wire: DocumentOp[] = JSON.parse(JSON.stringify(applied.inverse));
  const undone = applyDocumentOps(applied.document, wire).unwrap();
  expect(undone.document).toStrictEqual(document);
  const redo: DocumentOp[] = JSON.parse(JSON.stringify(undone.inverse));
  expect(
    applyDocumentOps(undone.document, redo).unwrap().document,
  ).toStrictEqual(applied.document);
  expect(document).toStrictEqual(original);
  return applied;
};

test(
  "note cleanup composes comment retirement and lifecycle inverses in reverse order",
  () => {
    assertProperty(
      fc.property(textArbitrary, (text) => {
        for (const kind of ["footnote", "endnote"] as const) {
          const original = sourceDocument(text);
          const ops: DocumentOp[] = [];
          let document = original;
          const apply = (op: DocumentOp) => {
            ops.push(op);
            document = applyDocumentOp(document, op).unwrap().document;
          };
          const at = { story: "main", blockId: "10000001", offset: 0 } as const;
          apply({
            type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
            comment: {
              id: 7,
              author: "Reviewer",
              done: false,
              content: [paragraph("20000007", text)],
            },
            anchor: {
              kind: "range",
              from: at,
              to: { ...at, offset: text.length },
            },
          });
          apply({
            type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
            comment: {
              id: 8,
              author: "Reviewer",
              content: [paragraph("20000008", "reply")],
            },
            anchor: { kind: "reply", parentId: 7 },
          });
          apply({
            type: DOCUMENT_OP_TYPES.ADD_NOTE,
            at: { ...at, offset: text.length },
            note: { type: kind, id: 2, content: [paragraph("20000002", text)] },
          });
          apply({
            type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
            comment: {
              id: 9,
              author: "Reviewer",
              done: false,
              content: [paragraph("20000009", "note comment")],
            },
            anchor: {
              kind: "point",
              at: { story: { kind, id: 2 }, blockId: "20000002", offset: 1 },
            },
          });
          apply({
            type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
            comment: {
              id: 10,
              author: "Reviewer",
              content: [paragraph("2000000A", "nested reply")],
            },
            anchor: { kind: "reply", parentId: 9 },
          });
          const beforeDeletion = document;
          const deletion = {
            type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
            story: "main",
            blockIds: [at.blockId],
          } as const satisfies DocumentOp;
          expectRestoration(beforeDeletion, [deletion]);
          apply(deletion);
          const refused = applyDocumentOp(beforeDeletion, {
            type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
            story: "main",
            blockIds: [at.blockId],
            revision,
          });
          expect(refused.isErr()).toBe(true);
          if (refused.isErr())
            expect(refused.error.reason).toBe(
              DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            );
          expect(
            document.package[kind === "footnote" ? "footnotes" : "endnotes"],
          ).toStrictEqual(
            original.package[kind === "footnote" ? "footnotes" : "endnotes"],
          );
          for (const key of ["footnotes", "endnotes"] as const)
            expect(document.package[key]?.at(0)).toBe(
              original.package[key]?.at(0),
            );
          expect(document.package.document.comments).toEqual(
            original.package.document.comments,
          );
          expect(
            [...(document.package.relationships?.values() ?? [])].some(
              ({ type }) =>
                type === COMMENT_PART_RELATIONSHIPS.commentsExtended.type,
            ),
          ).toBe(false);
          const composed = expectRestoration(original, ops);
          expect(composed.document).toStrictEqual(document);
          expect(document.package.document.content).toStrictEqual([
            original.package.document.content.at(1),
          ]);
        }
      }),
      {
        id: "note cleanup composes comment retirement and lifecycle inverses in reverse order",
        numRuns: 20,
      },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "undoing a first reference preserves authored orphan note declarations and their comment threads",
  () => {
    assertProperty(
      fc.property(textArbitrary, (text) => {
        for (const kind of ["footnote", "endnote"] as const) {
          for (const mode of ["direct", "tracked"] as const) {
            const original = sourceDocument(text);
            const note = {
              type: kind,
              id: 2,
              content: [paragraph("20000002", text)],
            };
            if (kind === "footnote")
              original.package.footnotes?.push({ ...note, type: kind });
            else original.package.endnotes?.push({ ...note, type: kind });
            const comment = {
              type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
              comment: {
                id: 7,
                author: "Reviewer",
                done: false,
                content: [paragraph("20000007", "authored note comment")],
              },
              anchor: {
                kind: "point",
                at: { story: { kind, id: 2 }, blockId: "20000002", offset: 1 },
              },
            } as const satisfies DocumentOp;
            const commented = applyDocumentOp(original, comment).unwrap()
              .document;
            const insert = {
              type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
              at: { story: "main", blockId: "10000001", offset: 1 },
              slice: {
                content: [
                  {
                    type: "run",
                    content: [
                      {
                        type:
                          kind === "footnote" ? "footnoteRef" : "endnoteRef",
                        id: 2,
                      },
                    ],
                  },
                ],
                openStart: 0,
                openEnd: 0,
              },
              ...(mode === "tracked" ? { revision } : {}),
            } as const satisfies DocumentOp;
            const applied = expectRestoration(commented, [insert]);
            expect(applied.document.package.document.comments).toEqual(
              commented.package.document.comments,
            );
            expectRestoration(original, [comment, insert]);
          }
        }
      }),
      {
        id: "undoing a first reference preserves authored orphan note declarations and their comment threads",
        numRuns: 20,
      },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "note-reference swaps inspect multi-operation inverses without recursively augmenting them",
  () => {
    assertProperty(
      fc.property(textArbitrary, (text) => {
        for (const kind of ["footnote", "endnote"] as const) {
          const original = sourceDocument(text);
          const orphan = {
            type: kind,
            id: 2,
            content: [paragraph("20000002", "authored orphan")],
          };
          if (kind === "footnote")
            original.package.footnotes?.push({ ...orphan, type: kind });
          else original.package.endnotes?.push({ ...orphan, type: kind });
          const noted = applyDocumentOp(original, {
            type: DOCUMENT_OP_TYPES.ADD_NOTE,
            at: { story: "main", blockId: "10000001", offset: 1 },
            note: {
              type: kind,
              id: 1,
              content: [paragraph("20000001", "referenced note")],
            },
          }).unwrap().document;
          const block = noted.package.document.content.at(0);
          expect(block?.type).toBe("paragraph");
          if (block?.type !== "paragraph")
            throw new TypeError("Missing note reference host.");
          const applied = expectRestoration(noted, [
            {
              type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
              story: "main",
              blockId: "10000001",
              expected: block.content,
              content: [
                {
                  type: "run",
                  content: [
                    { type: "text", text },
                    {
                      type: kind === "footnote" ? "footnoteRef" : "endnoteRef",
                      id: 2,
                    },
                  ],
                },
              ],
            },
          ]);
          expect(applied.inverse.map(({ type }) => type)).toEqual([
            DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
            DOCUMENT_OP_TYPES.REPLACE_INLINE,
            DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
          ]);
          expect(
            applied.document.package[
              kind === "footnote" ? "footnotes" : "endnotes"
            ]?.map(({ id }) => id),
          ).toEqual([3, 2]);
          expectRestoration(noted, [
            {
              type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
              story: "main",
              at: { type: "after", blockId: "10000001" },
              blocks: [
                {
                  type: "paragraph",
                  paraId: "40000001",
                  content: [
                    {
                      type: "run",
                      content: [
                        {
                          type:
                            kind === "footnote" ? "footnoteRef" : "endnoteRef",
                          id: 2,
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ]);
        }
      }),
      {
        id: "note-reference swaps inspect multi-operation inverses without recursively augmenting them",
        numRuns: 20,
      },
    );
  },
  propertyTestTimeout(30_000),
);
