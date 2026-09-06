/**
 * The snapshot walk and the live-document walk have to agree.
 *
 * A snapshot hands out one id per block. Applying an operation resolves that
 * id back to a position by walking the live document, and the two walks are
 * separate code. When they disagree about which nodes count — a blank
 * paragraph, a hidden row's subtree, a paragraph inside a nested table — every
 * id after the disagreement resolves onto the wrong block, and the edit lands
 * somewhere nobody asked for. That is not a failure any single example finds
 * reliably, so it is stated here as a property over generated documents.
 *
 * The observable form of the invariant: give every block of a snapshot its own
 * distinct style, in one batch, addressed by id. Every block must come back
 * carrying the style meant for it, in the same order, in the same container.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  buildBodySequenceDocx,
  type BodyItem,
  type CellContent,
} from "../compare/__fixtures__/body-sequence";
import type { FolioAIBlock } from "./types";
import { FolioDocxReviewer } from "./headless";

const wordArb = fc.stringMatching(/^[A-Za-z]{3,9}$/u);

/** Blank paragraphs are ordinary blocks, so a generated document holds them. */
const textArb = fc.oneof(
  { weight: 4, arbitrary: wordArb },
  { weight: 1, arbitrary: fc.constant("") },
);

const paragraphArb: fc.Arbitrary<BodyItem> = textArb.map((text) => ({
  kind: "paragraph",
  text,
}));

const nestedCellArb: fc.Arbitrary<CellContent> = fc.oneof(
  { weight: 5, arbitrary: textArb },
  {
    weight: 1,
    arbitrary: fc.array(textArb, { minLength: 1, maxLength: 2 }).map(
      (texts): CellContent => [
        { kind: "paragraph", text: texts[0] ?? "" },
        // A table as a cell's last child is the container edge the fixture
        // closes with the paragraph the format requires.
        { kind: "table", rows: [[texts[1] ?? ""]] },
      ],
    ),
  },
);

const tableArb: fc.Arbitrary<BodyItem> = fc
  .record({
    rows: fc.array(fc.array(nestedCellArb, { minLength: 1, maxLength: 2 }), {
      minLength: 1,
      maxLength: 3,
    }),
    hiddenRow: fc.option(fc.nat({ max: 2 }), { nil: undefined }),
  })
  .map(({ rows, hiddenRow }) => {
    // Every row the table's full width: a ragged one is not a table any
    // package holds, and the parser would square it before the snapshot saw it.
    const width = Math.max(...rows.map((row) => row.length));
    const squared: CellContent[][] = [];
    for (const row of rows) {
      const cells: CellContent[] = [];
      cells.push(...row);
      while (cells.length < width) {
        cells.push("");
      }
      squared.push(cells);
    }
    // Never hide every row: a table with nothing visible is a different shape
    // from a table with a hidden row in it.
    const hides = hiddenRow !== undefined && hiddenRow < squared.length && squared.length > 1;
    const table: BodyItem = { kind: "table", rows: squared };
    if (hides) {
      table.hiddenRows = [hiddenRow];
    }
    return table;
  });

/**
 * A body always ends with a paragraph: a table may not be a body's or a cell's
 * last child, and a document that breaks that rule is malformed input rather
 * than a case this property is about.
 */
const bodyArb: fc.Arbitrary<readonly BodyItem[]> = fc
  .array(fc.oneof({ weight: 3, arbitrary: paragraphArb }, { weight: 1, arbitrary: tableArb }), {
    minLength: 1,
    maxLength: 6,
  })
  .map((items) => {
    const body: BodyItem[] = [];
    body.push(...items);
    body.push({ kind: "paragraph", text: "closing" });
    return body;
  });

type BlockShape = { id: string; container: string; text: string };

const shapeOf = ({ id, text, table }: FolioAIBlock): BlockShape => ({
  id,
  text,
  container: table
    ? `t${String(table.tableIndex)}r${String(table.rowIndex)}c${String(table.cellIndex)}p${String(table.paragraphIndex)}`
    : "body",
});

/** A style id that names the block it was meant for, so a swap is visible. */
const styleFor = (index: number): string => `Probe${String(index).padStart(3, "0")}`;

describe("block ids resolve against the live document", () => {
  test(
    "every block of a snapshot resolves to the block it was taken from",
    async () => {
      await fc.assert(
        fc.asyncProperty(bodyArb, async (items) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(await buildBodySequenceDocx(items), {
            author: "probe",
          });
          const before = reviewer.snapshot().blocks;

          const { skipped } = reviewer.applyOperations(
            before.map((block, index) => ({
              id: `probe-${String(index)}`,
              type: "setBlockParagraphProperties" as const,
              blockId: block.id,
              properties: { styleId: styleFor(index) },
            })),
            { mode: "direct" },
          );
          expect(skipped).toEqual([]);

          const after = reviewer.snapshot().blocks;
          // The walk is the same walk: same blocks, same order, same
          // containers, same text — the batch changed only styles.
          expect(after.map(shapeOf)).toEqual(before.map(shapeOf));
          expect(after.map(({ styleId }) => styleId)).toEqual(before.map((_, i) => styleFor(i)));
        }),
        propertyConfig({ numRuns: 25 }),
      );
    },
    propertyTestTimeout(120_000),
  );
});
