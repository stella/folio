/**
 * A nested table cuts its cell's paragraphs into runs.
 *
 * The paragraphs before a cell's nested table and the ones after it are two
 * sequences, not one: nothing moves a block from one side of a table to the
 * other, and the last paragraph of a cell cannot be deleted because its mark
 * has nothing to join. Aligned as a single sequence, a shrinking cell pairs a
 * surviving paragraph with one on the far side of its nested table, and the
 * redline then keeps a paragraph the target does not have.
 *
 * The same cells reach a comparison with their closing paragraph missing,
 * because a producer that rewrites one deletes the whole run of paragraphs
 * including the one after the nested table. `CT_Tc` still ends in a paragraph,
 * so that package states a cell no consumer renders as written, and reading it
 * as written asks the comparison to delete a paragraph mark that has to stay.
 */

import { expect, test } from "bun:test";

import { buildBodySequenceDocx, type BodyItem, type Cell } from "./__fixtures__/body-sequence";
import { compareDocx } from "./compare";

const OPTIONS = {
  author: "compare",
  timestamp: "2024-03-01T00:00:00.000Z",
  onUnverified: "emit",
} as const;

const INTRO = { kind: "paragraph", text: "This agreement is made between the parties." } as const;
const OUTRO = { kind: "paragraph", text: "Signed by the parties above." } as const;

const SCHEDULE = {
  kind: "table",
  columnWidths: [1200, 1200],
  rows: [[{ content: "Each occurrence" }, { content: "Aggregate" }]],
} as const satisfies BodyItem;

const paragraph = (text: string) => ({ kind: "paragraph", text }) as const;

type CoverCellOptions = {
  /** Paragraphs before the nested schedule. */
  notes: readonly string[];
  /** The paragraph after it, or `null` for a package that omits it. */
  closing: string | null;
};

/** The value cell: `notes`, the nested schedule, then the closing paragraph. */
const coverCell = ({ notes, closing }: CoverCellOptions): Cell => ({
  content: [...notes.map(paragraph), SCHEDULE, ...(closing === null ? [] : [paragraph(closing)])],
  conformance: closing === null ? "unclosed" : "conforming",
});

const coverTable = (options: CoverCellOptions): BodyItem => ({
  kind: "table",
  columnWidths: [2400, 4800],
  rows: [[{ content: "Insurance minimums" }, coverCell(options)]],
});

const BASE_NOTES = [
  "Provider will carry the policies below.",
  "",
  "[Drafting note: keep the lines that apply.]",
  "",
] as const;

const NEGOTIATED_NOTES = ["Provider will carry the policies below."] as const;

const verificationOf = async (base: ArrayBuffer, target: ArrayBuffer) => {
  const result = await compareDocx(base, target, OPTIONS);
  if (result.isErr()) {
    throw result.error;
  }
  return result.value.verification;
};

const comparing = async (base: CoverCellOptions, target: CoverCellOptions) =>
  await verificationOf(
    await buildBodySequenceDocx([INTRO, coverTable(base), OUTRO]),
    await buildBodySequenceDocx([INTRO, coverTable(target), OUTRO]),
  );

test("a cell whose paragraphs shrink around its nested table round-trips", async () => {
  expect(
    await comparing({ notes: BASE_NOTES, closing: "" }, { notes: NEGOTIATED_NOTES, closing: "" }),
  ).toEqual({ status: "verified" });
});

/**
 * The closing paragraph is the one a comparison cannot delete, so a target
 * that rewords it has to pair it with the base's own rather than with a
 * paragraph from before the nested table.
 */
test("a cell whose closing paragraph is reworded round-trips", async () => {
  expect(
    await comparing(
      { notes: BASE_NOTES, closing: "Certificates on request." },
      { notes: NEGOTIATED_NOTES, closing: "Certificates within ten days of request." },
    ),
  ).toEqual({ status: "verified" });
});

test("a target that omits the cell's closing paragraph round-trips", async () => {
  expect(
    await comparing({ notes: BASE_NOTES, closing: "" }, { notes: NEGOTIATED_NOTES, closing: null }),
  ).toEqual({ status: "verified" });
});
