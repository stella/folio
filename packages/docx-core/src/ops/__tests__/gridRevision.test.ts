import { expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import { MAX_REVISION_ID, type Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { validateOpsDocument } from "../contract";
import { IDENTITY_SPACES, identityKeysIn, slotKey } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const documentOf = (
  revisionId: number,
  columnWidths: readonly (number | undefined)[] = [undefined, 1000],
): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          formatting: { gridChange: { id: revisionId, columnWidths } },
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", paraId: "00000001", content: [] }],
                },
              ],
            },
          ],
        },
        { type: "paragraph", paraId: "00000002", content: [] },
      ],
    },
  },
});

test("accepting a grid revision clears its historical snapshot", () => {
  const result = applyDocumentOp(documentOf(100), {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [100],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  if (result.isErr()) throw result.error;
  const table = result.value.document.package.document.content.at(0);
  if (table?.type !== "table") throw new Error("The table remains present.");
  expect(table.formatting?.gridChange).toBeUndefined();
});

test.each([
  [[undefined, 1000], DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE],
  [[2400, 1200], DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH],
] as const)(
  "rejecting an unrepresentable or inconsistent previous grid refuses atomically",
  (widths, reason) => {
    const document = documentOf(100, widths);
    const before = structuredClone(document);
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [100],
      decision: REVISION_DECISIONS.REJECT,
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.reason).toBe(reason);
    expect(document).toStrictEqual(before);
  },
);

test(
  "the revision census reserves every grid revision id and detects collisions",
  () => {
    fc.assert(
      fc.property(fc.integer({ min: 0, max: MAX_REVISION_ID }), (id) => {
        const document = documentOf(id);
        expect(identityKeysIn(document.package.document.content)).toEqual([
          slotKey({ space: IDENTITY_SPACES.REVISION, id }),
        ]);
        const tracked = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
          text: "x",
          runProps: {},
          revision: { id, author: "Reviewer", date: "2026-05-06T07:08:09Z" },
        });
        expect(tracked.isErr()).toBe(true);
        if (tracked.isErr())
          expect(tracked.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION);
        const duplicate = documentOf(id);
        duplicate.package.document.content.push({
          type: "table",
          formatting: { gridChange: { id, columnWidths: [] } },
          rows: [],
        });
        const valid = validateOpsDocument(duplicate);
        expect(valid.isErr()).toBe(true);
        if (valid.isErr())
          expect(valid.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID);
      }),
      propertyConfig({ numRuns: 2_000 }),
    );
  },
  propertyTestTimeout(60_000),
);
