/**
 * The persisted form of operations in the current schema version.
 *
 * Journaled operations are read back years after they were written, so their
 * JSON is pinned: one envelope per operation kind and per inverse, produced
 * from a fixed document. A change to an operation's fields, or to a model
 * record an operation embeds, shows up here as a diff of the fixture; it
 * needs a new schema version and a migration, not an updated fixture.
 *
 * `bun packages/docx-core/src/ops/__tests__/wireFixtures.ts` generates the current fixture.
 */

import { expect, test } from "bun:test";
import type { Paragraph } from "../../model/document";
import { DOCUMENT_OP_SCHEMA_VERSION, DOCUMENT_OP_TYPES } from "../types";
import { envelopes, wireFixturePath } from "./wireFixtures";

/** Paragraph reviews have one OOXML record; run-property review arrays have a separate contract. */
const expectParagraphReviewCardinality = (value: unknown): void => {
  if (Array.isArray(value)) {
    for (const entry of value) expectParagraphReviewCardinality(entry);
    return;
  }
  if (value === null || typeof value !== "object") return;
  if ("type" in value && value.type === "paragraph" && "propertyChanges" in value) {
    expect(Array.isArray(value.propertyChanges)).toBe(true);
    if (Array.isArray(value.propertyChanges))
      expect(value.propertyChanges.length).toBeLessThanOrEqual(1);
  }
  for (const entry of Object.values(value)) expectParagraphReviewCardinality(entry);
};

test("the wire cardinality guard rejects two property reviews on an identified paragraph", () => {
  const valid = {
    type: "paragraph",
    paraId: "00000001",
    content: [],
    propertyChanges: [
      { type: "paragraphPropertyChange", info: { id: 7, author: "Original reviewer" } },
    ],
  } satisfies Paragraph;
  expect(() => expectParagraphReviewCardinality(valid)).not.toThrow();
  const invalid = {
    ...valid,
    propertyChanges: [
      ...valid.propertyChanges,
      { type: "paragraphPropertyChange", info: { id: 8, author: "Latest reviewer" } },
    ],
  } satisfies Paragraph;
  expect(() => expectParagraphReviewCardinality(invalid)).toThrow();
});

test("every operation kind and its inverse keep their persisted JSON", async () => {
  const written = JSON.stringify(envelopes(), null, 2);
  if (process.env["UPDATE_OP_WIRE_FIXTURE"] === "1") {
    await Bun.write(wireFixturePath, `${written}\n`);
  }
  // The JSON as persisted, compared as data: the fixture's layout is the formatter's.
  const pinned: unknown = await Bun.file(wireFixturePath).json();
  expectParagraphReviewCardinality(pinned);
  const generated = envelopes();
  expectParagraphReviewCardinality(generated);
  for (const { op } of generated) {
    if (op.type !== DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW) continue;
    expect(op.expected.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
    expect(op.review.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
  }
  expect(JSON.parse(written)).toStrictEqual(pinned);
  const kinds = new Set(envelopes().map(({ op }) => op.type));
  expect([...kinds].toSorted()).toEqual(Object.values(DOCUMENT_OP_TYPES).toSorted());
  expect(envelopes().every(({ schema }) => schema === DOCUMENT_OP_SCHEMA_VERSION)).toBe(true);
});
