import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const existingIds = fc.uniqueArray(
  fc.oneof(
    fc.integer({ min: 1, max: MAX_REVISION_ID }),
    fc.constant(MAX_REVISION_ID),
  ),
  { minLength: 1, maxLength: 4 },
);

describe("headless comment IDs", () => {
  test(
    "comment and revision batches avoid loaded IDs and retain IDs on round-trip",
    async () => {
      // Existing uniqueness fixtures lacked an upper-bound oracle and authored
      // high IDs. Generate both comment and revision spaces across the full range.
      await assertProperty(
        fc.asyncProperty(existingIds, async (ids) => {
          const document = createEmptyDocument();
          document.package.document.comments = ids.map((id) => ({
            id,
            author: "Author",
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [{ type: "text", text: "Existing" }],
                  },
                ],
              },
            ],
          }));
          document.package.document.content = [
            {
              type: "paragraph",
              paraId: "12345678",
              content: [
                ...ids.map((id) => ({
                  type: "commentRangeStart" as const,
                  id,
                })),
                {
                  type: "run",
                  content: [{ type: "text", text: "Review this clause." }],
                },
                ...ids.map((id) => ({ type: "commentRangeEnd" as const, id })),
                ...ids.map((id) => ({ type: "commentReference" as const, id })),
              ],
            },
            {
              type: "paragraph",
              paraId: "23456789",
              content: ids.map((id) => ({
                type: "insertion",
                info: { id, author: "Author", date: "2026-01-01T00:00:00Z" },
                content: [
                  {
                    type: "run",
                    content: [{ type: "text", text: "Existing revision" }],
                  },
                ],
              })),
            },
          ];
          const buffer = await createDocx(document);
          for (const session of SESSIONS) {
            // Comment ids belong to the document: a second reviewer making the
            // same edits to it allocates the same ones. Revision ids are unique
            // within the document a reviewer edits; a second reviewer edits its
            // own copy, so it may draw the ones the first drew.
            const commentIdsByReviewer: number[][] = [];
            for (let reviewerIndex = 0; reviewerIndex < 2; reviewerIndex += 1) {
              const issuedRevisions = new Set(ids);
              const reviewer = await FolioDocxReviewer.fromBuffer(buffer, {
                session,
              });
              const block =
                reviewer.snapshot().blocks.at(0) ??
                panic("Expected a reviewable block");
              for (let batchIndex = 0; batchIndex < 2; batchIndex += 1) {
                const result = reviewer.applyOperations(
                  [
                    {
                      id: `comment-${batchIndex}`,
                      type: "commentOnBlock",
                      blockId: block.id,
                      comment: { text: "New note" },
                    },
                    {
                      id: `replace-${batchIndex}`,
                      type: "replaceInBlock",
                      blockId: block.id,
                      find: batchIndex === 0 ? "Review" : "clause",
                      replace: batchIndex === 0 ? "Inspect" : "provision",
                    },
                  ],
                  { undefinedReferences: "refuse" },
                );
                expect(result.skipped).toEqual([]);
                expect(result.applied).toHaveLength(2);
                const revisionIds =
                  result.applied.find(
                    ({ id }) => id === `replace-${batchIndex}`,
                  )?.revisionIds ?? [];
                expect(revisionIds.length).toBeGreaterThan(0);
                for (const id of revisionIds) {
                  expect(Number.isInteger(id)).toBe(true);
                  expect(id).toBeGreaterThan(0);
                  expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
                  expect(issuedRevisions.has(id)).toBe(false);
                  issuedRevisions.add(id);
                }
              }
              const parentId = ids.at(0) ?? panic("Expected a loaded comment");
              for (let replyIndex = 0; replyIndex < 2; replyIndex += 1) {
                expect(
                  reviewer.replyTo(parentId, { text: "Reply note" }),
                ).not.toBeNull();
              }
              const newIds =
                reviewer
                  .toDocument()
                  .package.document.comments?.filter(
                    ({ author }) => author !== "Author",
                  )
                  .map(({ id }) => id) ?? [];
              expect(newIds).toHaveLength(4);
              const issued = new Set(ids);
              for (const id of newIds) {
                expect(Number.isInteger(id)).toBe(true);
                // A comment w:id is any non-negative decimal; once the highest
                // is taken, a fresh one wraps to the lowest free id, 0 included.
                expect(id).toBeGreaterThanOrEqual(0);
                expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
                expect(issued.has(id)).toBe(false);
                issued.add(id);
              }
              commentIdsByReviewer.push(newIds);
              const saved = await reviewer.toBuffer();
              const reopened = await parseDocx(saved, { preloadFonts: false });
              const reopenedReviewer = await FolioDocxReviewer.fromBuffer(
                saved,
                {
                  session,
                },
              );
              expect(
                new Set(reopenedReviewer.getChanges().map(({ id }) => id)),
              ).toEqual(new Set(reviewer.getChanges().map(({ id }) => id)));
              expect(
                new Set(
                  reopened.package.document.comments?.map(({ id }) => id),
                ),
              ).toEqual(new Set([...ids, ...newIds]));
            }
            expect(commentIdsByReviewer.at(1)).toEqual(
              commentIdsByReviewer.at(0),
            );
          }
        }),
        {
          numRuns: 12,
          id: "comment and revision batches avoid loaded IDs and retain IDs on round-trip",
        },
      );
    },
    propertyTestTimeout(30_000),
  );
});
