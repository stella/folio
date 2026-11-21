import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { FolioRevisionStamp } from "./apply";
import { FolioDocxReviewer } from "./headless";

const DELETION_STAMP = { date: "2026-01-01T00:00:00Z", idSeed: 100 };
const INSERTION_STAMP = { date: "2026-01-02T00:00:00Z", idSeed: 200 };

const INSERTION_PROPERTIES = {
  heading: { styleId: "Heading2" },
  bullet: { numbering: { kind: "newList", format: "bullet" } },
  numbered: { numbering: { kind: "newList", format: "numbered" } },
  paragraph: { alignment: "center", spacing: { spaceAfter: 240 } },
} as const;

// The canonical session refuses to move a pending formatting change onto a
// new final paragraph; legacy records the insertion.
const FINAL_MARK_REFUSAL =
  "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of review records.";

const shape = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map((block) => {
    if (block.kind === "diagnostic")
      panic("Generated prose fixture contains a diagnostic");
    const {
      kind,
      text,
      styleId,
      headingLevel,
      statedNumbering,
      directAlignment,
      directSpacing,
    } = block;
    return {
      kind,
      text,
      styleId,
      headingLevel,
      statedNumbering,
      directAlignment,
      directSpacing,
    };
  });

// The sequence generator previously sampled deletion and insertion independently.
// Cross every insertion property with a deleted terminal suffix on every run.
for (const engine of [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[])
  for (const finalKind of ["paragraph", "heading"] as const) {
    for (const removed of [1, 2, 3]) {
      for (const [kind, properties] of Object.entries(INSERTION_PROPERTIES)) {
        for (const session of ["live", "reopened"] as const) {
          test(`insert ${kind} after the last survivor of ${removed} terminal deletions (${finalKind}) / ${session} (${engine})`, async () => {
            // A heading's deletion leaves its pending style change on the
            // survivor's mark, which a tracked insertion would move onto the
            // new final paragraph.
            const refused = engine === "canonical" && finalKind === "heading";
            const bytes = await createDocx(
              fromMarkdown(
                Array.from(
                  { length: removed + 2 },
                  (_, index) =>
                    `${finalKind === "heading" && index === removed + 1 ? "## " : ""}Clause ${index}.`,
                ).join("\n\n"),
              ),
            );
            const open = (data: ArrayBuffer) =>
              FolioDocxReviewer.fromBuffer(data, { session: engine });
            const direct = await open(bytes);
            let tracked = await open(bytes);
            let rejected = await open(bytes);
            let selectivelyRejected = await open(bytes);
            const before = shape(tracked);
            const replay = (
              reviewer: FolioDocxReviewer,
              mode: FolioDocumentOperationMode,
            ) => {
              const blocks = reviewer.getContent();
              const apply = (
                operations: FolioDocumentOperation[],
                revisionStamp: FolioRevisionStamp,
              ) =>
                reviewer.applyDocumentOperations(
                  {
                    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                    mode,
                    operations,
                  },
                  { undefinedReferences: "refuse", revisionStamp },
                );
              const deletions = blocks
                .slice(2)
                .map(({ id }): FolioDocumentOperation => ({
                  id: `delete-${id}`,
                  type: "deleteBlock",
                  blockId: id,
                }));
              const deleted = apply(deletions, DELETION_STAMP);
              expect(deleted.skipped).toEqual([]);
              expect(deleted.applied).toHaveLength(deletions.length);
              // The last survivor as it now stands: applied directly, it joined
              // the deleted final paragraph and goes by that one's id.
              const anchor = reviewer.getContent().at(1);
              if (!anchor) panic("Generated fixture has no surviving anchor");
              const inserted = apply(
                [
                  {
                    id: "insert",
                    type: "insertAfterBlock",
                    blockId: anchor.id,
                    text: "Inserted clause.",
                    ...properties,
                  },
                ],
                INSERTION_STAMP,
              );
              if (refused && mode !== "direct") {
                expect(inserted.applied).toEqual([]);
                expect(inserted.skipped).toEqual([
                  expect.objectContaining({
                    id: "insert",
                    message: FINAL_MARK_REFUSAL,
                  }),
                ]);
                return;
              }
              expect(inserted.skipped).toEqual([]);
              expect(inserted.applied).toHaveLength(1);
            };
            replay(direct, "direct");
            replay(tracked, "tracked-changes");
            if (refused) return;
            replay(rejected, "tracked-changes");
            replay(selectivelyRejected, "tracked-changes");
            if (session === "reopened") {
              tracked = await open(await tracked.toBuffer());
              rejected = await open(await rejected.toBuffer());
              selectivelyRejected = await open(
                await selectivelyRejected.toBuffer(),
              );
            }
            expect(
              selectivelyRejected
                .getChanges()
                .filter(({ date }) => date === INSERTION_STAMP.date).length,
            ).toBeGreaterThan(0);
            for (const change of selectivelyRejected
              .getChanges()
              .filter(({ date }) => date === INSERTION_STAMP.date)) {
              selectivelyRejected.rejectChange(change);
            }
            expect(
              selectivelyRejected
                .getChanges()
                .some(({ date }) => date === INSERTION_STAMP.date),
            ).toBe(false);
            selectivelyRejected.acceptAll();
            expect(shape(selectivelyRejected)).toEqual(before.slice(0, 2));
            expect(
              shape(await open(await selectivelyRejected.toBuffer())),
            ).toEqual(before.slice(0, 2));
            tracked.acceptAll();
            expect(shape(tracked)).toEqual(shape(direct));
            rejected.rejectAll();
            expect(shape(rejected)).toEqual(before);
            expect(shape(await open(await tracked.toBuffer()))).toEqual(
              shape(direct),
            );
            expect(shape(await open(await rejected.toBuffer()))).toEqual(
              before,
            );
          });
        }
      }
    }
  }
