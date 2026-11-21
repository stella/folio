import { expect, test } from "bun:test";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  normalizeForOps,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { serializedOpParts } from "../../../../scripts/lib/corpus-invariants/op-sequences";
import { noteCommentDocument } from "../../../../test/generators/noteCommentDocument";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";

const revision = { id: 41, author: "Reviewer", date: "2026-11-17T00:00:00Z" };
// The note's text goes with its sole reference, under ids of its own.
const NOTE_IDS = {
  revision: Array.from({ length: 16 }, (_, index) => 42 + index),
};
for (const kind of ["footnote", "endnote"] as const)
  for (const anchor of ["point", "range"] as const) {
    test(`${kind} ${anchor} comment retirement survives parsed save, reopen and exact inverse`, async () => {
      const { document: authored, retiredIds } = noteCommentDocument({
        kind,
        anchor,
        text: "source note",
        depth: 2,
        references: "sole",
      });
      const original = normalizeForOps(
        await parseDocx(await createDocx(authored), { preloadFonts: false }),
      );
      const savedSource = await serializedOpParts(original);
      const at = {
        story: OP_STORIES.MAIN,
        blockId: "10000001",
        offset: "source note".length,
      } as const;
      for (const type of [
        DOCUMENT_OP_TYPES.DELETE_RANGE,
        DOCUMENT_OP_TYPES.REMOVE_NOTE,
      ])
        for (const mode of ["direct", "accept", "reject"] as const) {
          const deletion =
            type === DOCUMENT_OP_TYPES.DELETE_RANGE
              ? ({
                  type,
                  from: at,
                  to: { ...at, offset: at.offset + 1 },
                } as const)
              : ({ type, at, story: { kind, id: 1 } } as const);
          const ops: DocumentOp[] =
            mode === "direct"
              ? [deletion]
              : [
                  { ...deletion, revision, newIds: NOTE_IDS },
                  {
                    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                    story: OP_STORIES.MAIN,
                    revisionIds: [revision.id],
                    decision: mode,
                  },
                ];
          const applied = applyDocumentOps(original, ops).unwrap();
          const reopened = await parseDocx(
            await repackDocx(applied.document, { updateModifiedDate: false }),
            { preloadFonts: false },
          );
          const expectedComments =
            mode === "reject"
              ? original.package.document.comments
              : original.package.document.comments?.filter(
                  ({ id }) => !retiredIds.includes(id),
                );
          expect(
            reopened.package.document.comments?.map(({ id }) => id),
          ).toEqual(expectedComments?.map(({ id }) => id));
          expect(
            reopened.package[
              kind === "footnote" ? "footnotes" : "endnotes"
            ]?.some(({ id }) => id === 1),
          ).toBe(mode === "reject");
          expect(
            reopened.package.document.comments?.find(({ id }) => id === 99),
          ).toBeDefined();
          const inverse: DocumentOp[] = JSON.parse(
            JSON.stringify(applied.inverse),
          );
          const restored = applyDocumentOps(applied.document, inverse).unwrap();
          expect(restored.document).toStrictEqual(original);
          expect(await serializedOpParts(restored.document)).toEqual(
            savedSource,
          );
          const redo: DocumentOp[] = JSON.parse(
            JSON.stringify(restored.inverse),
          );
          expect(
            await serializedOpParts(
              applyDocumentOps(restored.document, redo).unwrap().document,
            ),
          ).toEqual(await serializedOpParts(applied.document));
        }
    });
  }
