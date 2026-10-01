/**
 * Revision-id minting must stay inside the range OOXML consumers accept.
 *
 * Regression cover for the counter being seeded from `Date.now()` (~1.8e12),
 * which serialized straight into `<w:ins w:id="1784…"/>` and made exported
 * files unreadable. See `MAX_REVISION_ID` for the range rationale.
 * Port of eigenpal/docx-editor#1093.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { Schema } from "prosemirror-model";
import fc from "fast-check";

import { MAX_REVISION_ID } from "@stll/docx-core/model";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";

import {
  claimRevisionIds,
  mintRevisionId,
  nextRevisionId,
  nextRevisionIdRange,
  revisionIdSeedAbove,
  RevisionIdAllocationError,
  seedRevisionIdsAbove,
  seedRevisionIdsFromDoc,
} from "./revisionIds";

setDefaultTimeout(propertyTestTimeout(10_000));

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { pPrMark: { default: null }, cellMarker: { default: null } },
      toDOM: () => ["p", 0],
    },
    text: { group: "inline" },
  },
  marks: {
    runPropertyChange: {
      attrs: { changes: { default: [] } },
      toDOM: () => ["span", 0],
    },
    insertion: {
      attrs: { revisionId: { default: 0 }, author: { default: "" }, date: { default: "" } },
      toDOM: () => ["ins", 0],
    },
  },
});

test("deterministic revision seeds prefer valid ids above loaded values and skip occupied rollover intervals", () => {
  expect(revisionIdSeedAbove([])).toBe(1);
  expect(revisionIdSeedAbove([0, 2, 40])).toBe(41);
  expect(revisionIdSeedAbove([Number.NaN, -1, MAX_REVISION_ID + 1, 8])).toBe(9);
  const occupied = [0, 1, 3, MAX_REVISION_ID];
  const seed = revisionIdSeedAbove(occupied);
  expect(seed).toBe(4);
  expect(revisionIdSeedAbove(occupied.toReversed())).toBe(seed);
  mintRevisionId();
  expect(revisionIdSeedAbove(occupied)).toBe(seed);
  expect(seed).toBeLessThanOrEqual(MAX_REVISION_ID);
});

describe("mintRevisionId", () => {
  test("mints ids inside the signed 32-bit range OOXML consumers accept", () => {
    for (let i = 0; i < 3; i++) {
      const id = mintRevisionId();
      expect(id).toBeGreaterThan(0);
      expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
    }
  });

  test("mints strictly increasing ids", () => {
    const first = mintRevisionId();
    const second = mintRevisionId();
    expect(second).toBeGreaterThan(first);
  });
});

describe("seedRevisionIdsAbove", () => {
  test("resumes numbering just above the document's existing max id", () => {
    // The counter is module state shared with every earlier test in the run,
    // so seed relative to where it stands rather than to a fixed id.
    const maxId = mintRevisionId() + 1_000_000;
    seedRevisionIdsAbove(maxId);
    expect(mintRevisionId()).toBe(maxId + 1);
  });

  test("never lowers the counter", () => {
    seedRevisionIdsAbove(2_000_000);
    seedRevisionIdsAbove(5);
    expect(mintRevisionId()).toBeGreaterThan(2_000_000);
  });

  test("ignores an out-of-range id from an untrusted file", () => {
    seedRevisionIdsAbove(9e18);
    expect(mintRevisionId()).toBeLessThanOrEqual(MAX_REVISION_ID);
  });

  test("ignores a malformed max id", () => {
    seedRevisionIdsAbove(Number.NaN);
    seedRevisionIdsAbove(Number.POSITIVE_INFINITY);
    expect(mintRevisionId()).toBeLessThanOrEqual(MAX_REVISION_ID);
  });
});

describe("seedRevisionIdsFromDoc", () => {
  test("seeds above an id carried by an inline mark", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("x", [schema.marks["insertion"]!.create({ revisionId: 3_000_000 })]),
      ]),
    ]);

    seedRevisionIdsFromDoc(doc);

    expect(mintRevisionId()).toBeGreaterThan(3_000_000);
  });

  test("seeds above an id carried by a paragraph-mark attr", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { pPrMark: { kind: "ins", info: { id: 4_000_000, author: "A" } } }, [
        schema.text("x"),
      ]),
    ]);

    seedRevisionIdsFromDoc(doc);

    expect(mintRevisionId()).toBeGreaterThan(4_000_000);
  });

  test("seeds above an id nested under a cell marker", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        { cellMarker: { kind: "ins", info: { revisionId: 5_000_000, author: "A", date: null } } },
        [schema.text("x")],
      ),
    ]);

    seedRevisionIdsFromDoc(doc);

    expect(mintRevisionId()).toBeGreaterThan(5_000_000);
  });
});

test("reserves every loaded run-property revision, including on rollover", () => {
  const changes = [
    { info: { id: 7_000_000, author: "Author" } },
    { info: { id: MAX_REVISION_ID, author: "Author" } },
  ];
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("x", [schema.marks["runPropertyChange"]!.create({ changes })]),
    ]),
  ]);
  seedRevisionIdsFromDoc(doc);
  const id = mintRevisionId();
  expect(id).not.toBe(7_000_000);
  expect(id).not.toBe(MAX_REVISION_ID);
  expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
});

describe("range top boundary", () => {
  test("wrap skips loaded and previously issued ids", () => {
    const issued = mintRevisionId();
    seedRevisionIdsAbove(MAX_REVISION_ID - 1);
    const boundaryId = mintRevisionId();
    expect(boundaryId).toBeLessThanOrEqual(MAX_REVISION_ID);
    expect(boundaryId).not.toBe(issued);
    seedRevisionIdsAbove(MAX_REVISION_ID);
    const wrapped = mintRevisionId();
    expect(wrapped).toBeGreaterThan(0);
    expect(wrapped).not.toBe(issued);
    expect(wrapped).not.toBe(MAX_REVISION_ID);
  });

  test("rejects an overflowing batch before reserving it", () => {
    const before = nextRevisionId();
    expect(() => claimRevisionIds(MAX_REVISION_ID, MAX_REVISION_ID + 2)).toThrow(
      RevisionIdAllocationError,
    );
    expect(nextRevisionId()).toBe(before);
  });
});

test("shared contiguous batches avoid every occupied hole and high tail across producer schedules", () => {
  // The prior allocation oracle checked single ids. Generate multi-id demands
  // and loaded boundaries between batches, including single-id producers.
  assertProperty(
    fc.property(
      fc.array(
        fc.record({
          offset: fc.integer({ min: 0, max: 64 }),
          demand: fc.integer({ min: 2, max: 16 }),
        }),
        { minLength: 2, maxLength: 20 },
      ),
      (schedule) => {
        const loaded = new Set(schedule.map(({ offset }) => MAX_REVISION_ID - offset));
        for (const id of loaded) seedRevisionIdsAbove(id);
        const issued = new Set<number>();
        for (const { offset, demand } of schedule) {
          seedRevisionIdsAbove(MAX_REVISION_ID - offset);
          const first = nextRevisionIdRange();
          claimRevisionIds(first, first + demand);
          for (let id = first; id < first + demand; id += 1) {
            expect(id).toBeGreaterThan(0);
            expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
            expect(loaded.has(id)).toBe(false);
            expect(issued.has(id)).toBe(false);
            issued.add(id);
          }
          const reentrant = mintRevisionId();
          expect(loaded.has(reentrant)).toBe(false);
          expect(issued.has(reentrant)).toBe(false);
          issued.add(reentrant);
        }
      },
    ),
    { numRuns: 40 },
  );
});
