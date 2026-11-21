import { expect, test } from "bun:test";
import fc from "fast-check";

import { noteCommentDocument } from "../../../../../test/generators/noteCommentDocument";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import {
  COMMENT_PART_RELATIONSHIPS,
  type Document,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const revision = { id: 41, author: "Reviewer", date: "2026-11-17T00:00:00Z" };
const NOTE_IDS = {
  revision: Array.from({ length: 16 }, (_, index) => 42 + index),
};
const expectExact = (document: Document, ops: readonly DocumentOp[]) => {
  const original = structuredClone(document);
  const edit = applyDocumentOps(document, ops).unwrap();
  expect(contractViolation(edit.document)).toBeUndefined();
  const inverse: DocumentOp[] = JSON.parse(JSON.stringify(edit.inverse));
  const undone = applyDocumentOps(edit.document, inverse).unwrap();
  expect(undone.document).toStrictEqual(document);
  const redo: DocumentOp[] = JSON.parse(JSON.stringify(undone.inverse));
  expect(
    applyDocumentOps(undone.document, redo).unwrap().document,
  ).toStrictEqual(edit.document);
  expect(document).toStrictEqual(original);
  return edit.document;
};

test(
  "note-body retirement owns only newly exhausted threads through direct and tracked review",
  () => {
    assertProperty(
      fc.property(
        fc
          .array(fc.constantFrom("a", "b", " "), { minLength: 2, maxLength: 8 })
          .map((letters) => letters.join("")),
        fc.integer({ min: 0, max: 3 }),
        (text, depth) => {
          for (const kind of ["footnote", "endnote"] as const)
            for (const anchor of ["point", "range"] as const) {
              const { document, retiredIds, otherKind } = noteCommentDocument({
                kind,
                anchor,
                text,
                depth,
                references: "sole",
              });
              const at = {
                story: OP_STORIES.MAIN,
                blockId: "10000001",
                offset: text.length,
              } as const;
              for (const type of [
                DOCUMENT_OP_TYPES.DELETE_RANGE,
                DOCUMENT_OP_TYPES.REMOVE_NOTE,
              ]) {
                const direct =
                  type === DOCUMENT_OP_TYPES.DELETE_RANGE
                    ? ({
                        type,
                        from: at,
                        to: { ...at, offset: text.length + 1 },
                      } as const)
                    : ({ type, at, story: { kind, id: 1 } } as const);
                const removed = expectExact(document, [direct]);
                expect(
                  removed.package[
                    kind === "footnote" ? "footnotes" : "endnotes"
                  ],
                ).toEqual([]);
                expect(removed.package.document.comments).toEqual(
                  document.package.document.comments?.filter(
                    ({ id }) => !retiredIds.includes(id),
                  ),
                );
                expect(
                  [...(removed.package.relationships?.values() ?? [])].some(
                    ({ type: relationshipType }) =>
                      relationshipType ===
                      COMMENT_PART_RELATIONSHIPS.commentsExtended.type,
                  ),
                ).toBe(false);
                expect(
                  removed.package[
                    otherKind === "footnote" ? "footnotes" : "endnotes"
                  ]?.at(0),
                ).toBe(
                  document.package[
                    otherKind === "footnote" ? "footnotes" : "endnotes"
                  ]?.at(0),
                );
                // The note's text goes with its sole reference, under ids of its own.
                const tracked = { ...direct, revision, newIds: NOTE_IDS };
                const pending = expectExact(document, [tracked]);
                expect(pending.package.document.comments).toEqual(
                  document.package.document.comments,
                );
                for (const decision of ["accept", "reject"] as const) {
                  const resolve = {
                    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                    story: OP_STORIES.MAIN,
                    revisionIds: [revision.id],
                    decision,
                  } as const;
                  const resolved = expectExact(pending, [resolve]);
                  expect(resolved).toStrictEqual(
                    decision === "accept" ? removed : document,
                  );
                  expect(
                    expectExact(document, [tracked, resolve]),
                  ).toStrictEqual(resolved);
                }
              }
              const shared = noteCommentDocument({
                kind,
                anchor,
                text,
                depth,
                references: "shared",
              }).document;
              for (const mode of ["direct", "tracked"] as const) {
                const deletion = {
                  type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                  from: at,
                  to: { ...at, offset: text.length + 1 },
                  ...(mode === "tracked" ? { revision } : {}),
                } as const;
                let retained = expectExact(shared, [deletion]);
                if (mode === "tracked")
                  retained = expectExact(retained, [
                    {
                      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                      story: OP_STORIES.MAIN,
                      revisionIds: [revision.id],
                      decision: "accept",
                    },
                  ]);
                expect(retained.package.document.comments).toEqual(
                  shared.package.document.comments,
                );
                expect(
                  retained.package[
                    kind === "footnote" ? "footnotes" : "endnotes"
                  ],
                ).toEqual(
                  shared.package[
                    kind === "footnote" ? "footnotes" : "endnotes"
                  ],
                );
                const refused = applyDocumentOp(shared, {
                  type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
                  at,
                  story: { kind, id: 1 },
                  ...(mode === "tracked" ? { revision } : {}),
                });
                expect(refused.isErr()).toBe(true);
                if (refused.isErr())
                  expect(refused.error.reason).toBe(
                    DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
                  );
              }
            }
        },
      ),
      {
        id: "note-body retirement owns only newly exhausted threads through direct and tracked review",
        numRuns: 10,
        examples: [["aa", 3]],
      },
    );
  },
  propertyTestTimeout(30_000),
);
