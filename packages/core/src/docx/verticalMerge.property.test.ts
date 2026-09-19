/**
 * A vertical merge keeps its origin through the editor.
 *
 * `w:vMerge="restart"` is carried across the ProseMirror model by the cell's
 * rowspan, and a rowspan only exists once a continuation cell joins the
 * origin. Every column whose merge closed with a rowspan of one therefore had
 * nothing left to say the cell was a merge origin, and `fromProseDoc` deleted
 * the `w:vMerge` it found in `_originalFormatting`: a restart with no
 * continuation at all, one a plain cell interrupts, one whose continuation a
 * tracked revision carries. Dropping it changes the table's visible structure.
 *
 * The property runs every column pattern of `restart` / `continue` / none over
 * two to four rows, with and without a `w:cellMerge` revision on the cells,
 * through both paths that rebuild the table:
 *
 *   save    parse → repack → parse
 *   editor  parse → toProseDoc → fromProseDoc → repack → parse
 *
 * and demands the column's merge states back unchanged.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Table, TableCell, TableCellFormatting } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

/** The three states a cell's `w:vMerge` can be in. */
const MERGE_STATES = ["none", "restart", "continue"] as const;
type MergeState = (typeof MERGE_STATES)[number];

/** Whether the cells also carry a `w:cellMerge` revision. */
const REVISION_STATES = ["untracked", "tracked"] as const;
type RevisionState = (typeof REVISION_STATES)[number];

type MergeCase = {
  column: readonly MergeState[];
  revision: RevisionState;
};

const TRACKED_CHANGE_INFO = {
  id: 1,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
} as const;

const vMergeOf = (state: MergeState): TableCellFormatting["vMerge"] =>
  state === "none" ? undefined : state;

const cellFor = (state: MergeState, revision: RevisionState, text: string): TableCell => {
  const vMerge = vMergeOf(state);
  return {
    type: "tableCell",
    ...(vMerge !== undefined || revision === "tracked"
      ? {
          formatting: {
            ...(vMerge !== undefined ? { vMerge } : {}),
            ...(revision === "tracked"
              ? {
                  structuralChange: {
                    type: "tableCellMerge" as const,
                    info: TRACKED_CHANGE_INFO,
                    verticalMerge: "continue" as const,
                  },
                }
              : {}),
          },
        }
      : {}),
    content: [{ type: "paragraph", content: [{ type: "run", content: [{ type: "text", text }] }] }],
  };
};

/**
 * One merged column beside a plain one, so the table stays a grid whatever
 * the generated column does.
 */
const tableFor = ({ column, revision }: MergeCase): Table => ({
  type: "table",
  rows: column.map((state, index) => ({
    type: "tableRow",
    cells: [
      cellFor(state, revision, `m${String(index)}`),
      cellFor("none", "untracked", `p${String(index)}`),
    ],
  })),
});

const mergeCase = fc.record({
  column: fc.array(fc.constantFrom(...MERGE_STATES), { minLength: 2, maxLength: 4 }),
  revision: fc.constantFrom(...REVISION_STATES),
});

const withTable = (document: Document, table: Table): Document => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content: [table] },
  },
});

/** The first column's merge state, row by row. */
const readColumn = (document: Document): (TableCellFormatting["vMerge"] | "missing")[] => {
  const block = document.package.document.content.at(0);
  if (!block || block.type !== "table") {
    return ["missing"];
  }
  return block.rows.map((row) => row.cells.at(0)?.formatting?.vMerge);
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { detectVariables: false, preloadFonts: false });

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

describe("a vertical merge keeps its origin", () => {
  test("a restart with no continuation is still a restart after an edit", async () => {
    const template = await parse(await createEmptyDocx());
    const opened = await parse(
      await save(withTable(template, tableFor({ column: ["restart"], revision: "untracked" }))),
    );
    expect(readColumn(opened)).toEqual(["restart"]);

    const edited = fromProseDoc(toProseDoc(opened), opened);
    expect(readColumn(await parse(await save(edited)))).toEqual(["restart"]);
  });

  test(
    "every column pattern survives both rebuild paths, tracked or not",
    async () => {
      const template = await parse(await createEmptyDocx());

      await fc.assert(
        fc.asyncProperty(mergeCase, async (merge) => {
          const opened = await parse(await save(withTable(template, tableFor(merge))));
          const expected = readColumn(opened);
          expect(expected).toEqual(merge.column.map((state) => vMergeOf(state)));

          expect(readColumn(await parse(await save(opened)))).toEqual(expected);

          const edited = fromProseDoc(toProseDoc(opened), opened);
          expect(readColumn(await parse(await save(edited)))).toEqual(expected);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(90_000),
  );
});
