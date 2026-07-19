/**
 * Revision-id minting must stay inside the range OOXML consumers accept.
 *
 * Regression cover for the counter being seeded from `Date.now()` (~1.8e12),
 * which serialized straight into `<w:ins w:id="1784…"/>` and made exported
 * files unreadable. See `MAX_REVISION_ID` for the range rationale.
 * Port of eigenpal/docx-editor#1093.
 */

import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import { mintRevisionId, seedRevisionIdsAbove, seedRevisionIdsFromDoc } from "./revisionIds";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { pPrIns: { default: null }, cellMarker: { default: null } },
      toDOM: () => ["p", 0],
    },
    text: { group: "inline" },
  },
  marks: {
    insertion: {
      attrs: { revisionId: { default: 0 }, author: { default: "" }, date: { default: "" } },
      toDOM: () => ["ins", 0],
    },
  },
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
    seedRevisionIdsAbove(1_000_000);
    expect(mintRevisionId()).toBe(1_000_001);
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

  test("seeds above an id carried by a node attr", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { pPrIns: { revisionId: 4_000_000, author: "A", date: null } }, [
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

// MUST run last in this file: these drive the module counter to the very top of
// the range, so any later test in this file that expects a specific low value
// would see a wrapped counter.
describe("range top boundary", () => {
  test("seeding at the max id does not push the counter past the range", () => {
    seedRevisionIdsAbove(MAX_REVISION_ID - 1);
    seedRevisionIdsAbove(MAX_REVISION_ID);
    expect(mintRevisionId()).toBe(MAX_REVISION_ID);
  });
});
