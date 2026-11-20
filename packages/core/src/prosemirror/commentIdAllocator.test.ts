import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CommentIdAllocator,
  createCommentIdAllocator,
  seedCommentAllocator,
} from "./commentIdAllocator";
import { schema } from "./schema";

const loadedIds = fc.uniqueArray(
  fc.oneof(
    fc.integer({ min: 0, max: MAX_REVISION_ID }),
    fc.constant(MAX_REVISION_ID),
  ),
  { maxLength: 12 },
);

describe("comment ID allocation", () => {
  test(
    "a document's IDs stay in range, never repeat a loaded or issued ID, and ignore other documents",
    () => {
      // Vary loaded IDs across the full range, including the wrap boundary, and
      // what other documents in the process allocated first, up to the maximum.
      fc.assert(
        fc.property(loadedIds, loadedIds, (ids, otherIds) => {
          const issue = (allocator: CommentIdAllocator) => {
            for (const id of ids) allocator.seedAbove(id);
            return Array.from({ length: 24 }, () => allocator.next());
          };
          const alone = issue(createCommentIdAllocator());

          const other = createCommentIdAllocator();
          for (const id of otherIds) other.seedAbove(id);
          other.seedAbove(MAX_REVISION_ID - 1);
          for (let index = 0; index < 4; index += 1) other.next();
          expect(issue(createCommentIdAllocator())).toEqual(alone);

          const issued = new Set(ids);
          for (const id of alone) {
            expect(Number.isInteger(id)).toBe(true);
            expect(id).toBeGreaterThan(0);
            expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
            expect(issued.has(id)).toBe(false);
            issued.add(id);
          }
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(30_000),
  );

  test("a copy continues from its document's reservations, apart from the original", () => {
    const original = createCommentIdAllocator();
    original.seedAbove(7);
    const copy = original.clone();
    expect([original.next(), original.next()]).toEqual([8, 9]);
    expect(copy.next()).toBe(8);
  });

  test("reserves every live anchor even when comment metadata is absent", () => {
    const allocator = createCommentIdAllocator();
    const referenceId = allocator.next() + 20;
    const rangeId = referenceId + 1;
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("Marked", [
          (
            schema.marks["comment"] ?? panic("Expected the comment mark")
          ).create({
            commentId: MAX_REVISION_ID,
          }),
        ]),
        schema.node("commentReference", { commentId: referenceId }),
        schema.node("rangeAnchor", {
          markers: [
            { type: "commentRangeStart", id: rangeId },
            { type: "commentRangeEnd", id: rangeId },
          ],
        }),
      ]),
    ]);
    seedCommentAllocator(allocator, undefined, {
      state: EditorState.create({ doc }),
    });
    const reserved = new Set([MAX_REVISION_ID, referenceId, rangeId]);
    for (let index = 0; index < 48; index += 1) {
      expect(reserved.has(allocator.next())).toBe(false);
    }
  });

  test("untrusted out-of-range seeds cannot overflow the counter", () => {
    const allocator = createCommentIdAllocator();
    for (const id of [
      MAX_REVISION_ID + 1,
      Number.MAX_SAFE_INTEGER,
      -1,
      1.5,
      Infinity,
      NaN,
    ]) {
      allocator.seedAbove(id);
    }
    expect(allocator.next()).toBeLessThanOrEqual(MAX_REVISION_ID);
  });
});
