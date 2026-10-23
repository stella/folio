import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
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
for (const finalKind of ["paragraph", "heading"] as const) {
  for (const removed of [1, 2, 3]) {
    for (const [kind, properties] of Object.entries(INSERTION_PROPERTIES)) {
      for (const session of ["live", "reopened"] as const) {
        test(`insert ${kind} after the last survivor of ${removed} terminal deletions (${finalKind}) / ${session}`, async () => {
          const bytes = await createDocx(
            fromMarkdown(
              Array.from(
                { length: removed + 2 },
                (_, index) =>
                  `${finalKind === "heading" && index === removed + 1 ? "## " : ""}Clause ${index}.`,
              ).join("\n\n"),
            ),
          );
          const direct = await FolioDocxReviewer.fromBuffer(bytes);
          let tracked = await FolioDocxReviewer.fromBuffer(bytes);
          let rejected = await FolioDocxReviewer.fromBuffer(bytes);
          let selectivelyRejected = await FolioDocxReviewer.fromBuffer(bytes);
          const before = shape(tracked);
          const replay = (
            reviewer: FolioDocxReviewer,
            mode: FolioDocumentOperationMode,
          ) => {
            const blocks = reviewer.getContent();
            const anchor = blocks.at(1);
            if (!anchor) panic("Generated fixture has no surviving anchor");
            const apply = (
              operations: FolioDocumentOperation[],
              revisionStamp: FolioRevisionStamp,
            ) => {
              const result = reviewer.applyDocumentOperations(
                {
                  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                  mode,
                  operations,
                },
                { undefinedReferences: "refuse", revisionStamp },
              );
              expect(result.skipped).toEqual([]);
              expect(result.applied).toHaveLength(operations.length);
            };
            apply(
              blocks.slice(2).map(({ id }) => ({
                id: `delete-${id}`,
                type: "deleteBlock",
                blockId: id,
              })),
              DELETION_STAMP,
            );
            apply(
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
          };
          replay(direct, "direct");
          replay(tracked, "tracked-changes");
          replay(rejected, "tracked-changes");
          replay(selectivelyRejected, "tracked-changes");
          if (session === "reopened") {
            tracked = await FolioDocxReviewer.fromBuffer(
              await tracked.toBuffer(),
            );
            rejected = await FolioDocxReviewer.fromBuffer(
              await rejected.toBuffer(),
            );
            selectivelyRejected = await FolioDocxReviewer.fromBuffer(
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
            shape(
              await FolioDocxReviewer.fromBuffer(
                await selectivelyRejected.toBuffer(),
              ),
            ),
          ).toEqual(before.slice(0, 2));
          tracked.acceptAll();
          expect(shape(tracked)).toEqual(shape(direct));
          rejected.rejectAll();
          expect(shape(rejected)).toEqual(before);
          expect(
            shape(await FolioDocxReviewer.fromBuffer(await tracked.toBuffer())),
          ).toEqual(shape(direct));
          expect(
            shape(
              await FolioDocxReviewer.fromBuffer(await rejected.toBuffer()),
            ),
          ).toEqual(before);
        });
      }
    }
  }
}
