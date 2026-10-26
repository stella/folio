import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { Layout, Page } from "../layout-engine/types";
import type { HeadingInfo } from "./headingCollector";
import { resolveOutlineHeadings } from "./outlineHeadings";

/** Contiguous pages: page `i` holds `[starts[i], starts[i + 1])`. */
const layoutFromStarts = (starts: readonly number[], end: number): Layout => ({
  pageSize: { w: 816, h: 1056 },
  pages: starts.map((start, index): Page => ({
    number: index + 1,
    logicalNumber: index + 1,
    fragments: [
      {
        kind: "paragraph",
        blockId: `block-${String(index)}`,
        x: 72,
        y: 72,
        width: 600,
        height: 12,
        fromLine: 0,
        toLine: 1,
        pmStart: start,
        pmEnd: starts[index + 1] ?? end,
      },
    ],
    margins: { top: 72, right: 72, bottom: 72, left: 72 },
    size: { w: 816, h: 1056 },
  })),
});

const pageOf = (starts: readonly number[], pmPos: number) =>
  starts.findLastIndex((start) => start <= pmPos) + 1;

const DOCUMENT_END = 10_000;

const pageStarts = fc
  .uniqueArray(fc.integer({ min: 1, max: DOCUMENT_END - 1 }), {
    maxLength: 12,
  })
  .map((starts) => [0].concat(starts.sort((a, b) => a - b)));

const headingPositions = fc
  .uniqueArray(fc.integer({ min: 0, max: DOCUMENT_END - 1 }), {
    minLength: 1,
    maxLength: 30,
  })
  .map((positions) => positions.sort((a, b) => a - b));

/** What `collectHeadings` yields: fresh entries without page numbers. */
const collect = (positions: readonly number[]): HeadingInfo[] =>
  positions.map((pmPos, index) => ({
    id: `heading-${String(index)}`,
    text: `Heading ${String(index)}`,
    level: index % 3,
    pmPos,
  }));

describe("resolveOutlineHeadings", () => {
  test(
    "a reported page wins, an unreported one keeps the last known page, and only changed entries are replaced",
    () => {
      assertProperty(
        fc.property(
          headingPositions,
          pageStarts,
          pageStarts,
          (positions, before, after) => {
            const first = resolveOutlineHeadings({
              collected: collect(positions),
              layout: layoutFromStarts(before, DOCUMENT_END),
              previous: [],
            });
            expect(first.map(({ pageNumber }) => pageNumber)).toEqual(
              positions.map((pmPos) => pageOf(before, pmPos)),
            );

            // Settled: the same document and layout return the same list.
            expect(
              resolveOutlineHeadings({
                collected: collect(positions),
                layout: layoutFromStarts(before, DOCUMENT_END),
                previous: first,
              }),
            ).toBe(first);
            // Pending layout: nothing is blanked or replaced.
            expect(
              resolveOutlineHeadings({
                collected: collect(positions),
                layout: null,
                previous: first,
              }),
            ).toBe(first);

            const relaid = resolveOutlineHeadings({
              collected: collect(positions),
              layout: layoutFromStarts(after, DOCUMENT_END),
              previous: first,
            });
            for (const [index, heading] of relaid.entries()) {
              const previous = first[index];
              const page = pageOf(after, heading.pmPos);
              expect(heading.pageNumber).toBe(page);
              expect(heading === previous).toBe(previous?.pageNumber === page);
            }
          },
        ),
        {
          id: "a reported page wins, an unreported one keeps the last known page, and only changed entries are replaced",
        },
      );
    },
    propertyTestTimeout(5_000),
  );

  test("a heading the layout cannot place yet keeps its number under the same id", () => {
    const [kept, added] = resolveOutlineHeadings({
      collected: [
        { id: "kept", text: "Kept", level: 0, pmPos: 40 },
        { id: "added", text: "Added", level: 0, pmPos: 80 },
      ],
      layout: null,
      previous: [
        { id: "kept", text: "Kept", level: 0, pmPos: 12, pageNumber: 3 },
      ],
    });
    expect(kept?.pageNumber).toBe(3);
    expect(kept?.pmPos).toBe(40);
    expect(added?.pageNumber).toBeNull();
  });
});
