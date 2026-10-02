import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { MAX_REVISION_ID } from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "./schema";
import { createCommentIdAllocator, seedCommentAllocator } from "./commentIdAllocator";

const loadedIds = fc.uniqueArray(
  fc.oneof(fc.integer({ min: 0, max: MAX_REVISION_ID }), fc.constant(MAX_REVISION_ID)),
  { maxLength: 12 },
);

describe("comment ID allocation", () => {
  test(
    "all editor handles preserve the range and reserve loaded and issued IDs",
    () => {
      // The previous editor factory tests covered only small sequential IDs; vary
      // loaded IDs across the full range, including the wrap boundary.
      fc.assert(
        fc.property(loadedIds, (ids) => {
          const first = createCommentIdAllocator();
          const second = createCommentIdAllocator();
          for (const id of ids) first.seedAbove(id);
          const issued = new Set(ids);
          for (let index = 0; index < 24; index += 1) {
            const id = (index % 2 === 0 ? first : second).next();
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

  test("reserves every live anchor even when comment metadata is absent", () => {
    const allocator = createCommentIdAllocator();
    const referenceId = allocator.next() + 20;
    const rangeId = referenceId + 1;
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("Marked", [
          (schema.marks["comment"] ?? panic("Expected the comment mark")).create({
            commentId: MAX_REVISION_ID,
          }),
        ]),
        schema.node("commentReference", { commentId: referenceId }),
        schema.node("rangeAnchor", {
          start: { type: "commentRangeStart", id: rangeId },
          end: { type: "commentRangeEnd", id: rangeId },
        }),
      ]),
    ]);
    seedCommentAllocator(allocator, undefined, { state: EditorState.create({ doc }) });
    const reserved = new Set([MAX_REVISION_ID, referenceId, rangeId]);
    for (let index = 0; index < 48; index += 1) {
      expect(reserved.has(allocator.next())).toBe(false);
    }
  });

  test("untrusted out-of-range seeds cannot overflow the counter", () => {
    const allocator = createCommentIdAllocator();
    for (const id of [MAX_REVISION_ID + 1, Number.MAX_SAFE_INTEGER, -1, 1.5, Infinity, NaN]) {
      allocator.seedAbove(id);
    }
    expect(allocator.next()).toBeLessThanOrEqual(MAX_REVISION_ID);
  });
});
