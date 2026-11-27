/**
 * Column operations on a table whose file states no column widths, no grid at
 * all or grid columns without a measure, give in the canonical session the
 * package the default engine gives: pending, every change accepted and every
 * change rejected, in direct and tracked mode, and the canonical session's own
 * review of its saved package resolves them alike. A suggested column change
 * records the grid as the file states it, its columns without a measure.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { FolioDocumentOperation } from "../ai-edits";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import {
  type Mode,
  outcomes,
  packageOf,
  paragraph,
  reviewedInSession,
  row,
  table,
  textCell,
  viaCanonicalEngine,
  viaDefaultEngine,
} from "./__tests__/tableEngineParity";

setDefaultTimeout(propertyTestTimeout(240_000));

const rows = [
  row(textCell("20000001", "A1") + textCell("20000002", "B1")),
  row(textCell("20000003", "A2") + textCell("20000004", "B2")),
];

const TABLE_PROPERTIES = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>';

const BODIES = {
  "a table with no grid": `<w:tbl>${TABLE_PROPERTIES}${rows.join("")}</w:tbl>`,
  "a grid with no widths": `<w:tbl>${TABLE_PROPERTIES}<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${rows.join("")}</w:tbl>`,
} as const;
type Body = keyof typeof BODIES;

const ANCHORS = ["20000001", "20000002", "20000003", "20000004"] as const;

const OPERATIONS = {
  "insertTableColumn after": (blockId: string) => ({
    id: "op",
    type: "insertTableColumn",
    blockId,
    position: "after",
    cellTexts: ["New"],
  }),
  "insertTableColumn before": (blockId: string) => ({
    id: "op",
    type: "insertTableColumn",
    blockId,
    position: "before",
    cellTexts: ["New"],
  }),
  deleteTableColumn: (blockId: string) => ({
    id: "op",
    type: "deleteTableColumn",
    blockId,
  }),
} as const satisfies Record<
  string,
  (blockId: string) => FolioDocumentOperation
>;
type Operation = keyof typeof OPERATIONS;

const packageWith = (body: Body) =>
  packageOf([
    paragraph("10000001", "Before"),
    BODIES[body],
    paragraph("10000002", "After"),
  ]);

type Case = { body: Body; anchor: string; operation: Operation; mode: Mode };

type CanonicalRun = Awaited<ReturnType<typeof viaCanonicalEngine>>;

/** The case in both engines, and in the canonical session's own review. */
const sameAsDefaultEngine = async (
  { body, anchor, operation, mode }: Case,
  canonicalRun: (
    source: ArrayBuffer,
    op: FolioDocumentOperation,
  ) => Promise<CanonicalRun> = (source, op) =>
    viaCanonicalEngine(source, op, mode),
) => {
  const source = await packageWith(body);
  const op = OPERATIONS[operation](anchor);
  const legacy = await viaDefaultEngine(source, op, mode);
  const canonical = await canonicalRun(source, op);
  expect(canonical.refusals).toEqual([]);
  expect(canonical.skipped).toEqual(legacy.skipped);
  if (legacy.skipped.length > 0) return;
  const expected = await outcomes(legacy.bytes);
  expect(await outcomes(canonical.bytes)).toEqual(expected);
  expect(await reviewedInSession(canonical.bytes)).toEqual({
    accepted: expected.accepted,
    rejected: expected.rejected,
  });
};

const CASES = fc
  .constantFrom(...(Object.keys(BODIES) as Body[]))
  .chain((body) =>
    fc.record({
      body: fc.constant(body),
      anchor: fc.constantFrom(
        ...ANCHORS.filter((id) => BODIES[body].includes(`w14:paraId="${id}"`)),
      ),
      operation: fc.constantFrom(...(Object.keys(OPERATIONS) as Operation[])),
      mode: fc.constantFrom<Mode>("direct", "tracked-changes"),
    }),
  );

describe("column operations on a table whose file states no widths", () => {
  test("give the default engine's package in the canonical session, pending, accepted, rejected and resolved there", async () => {
    await assertProperty(
      fc.asyncProperty(CASES, (example) => sameAsDefaultEngine(example)),
      {
        numRuns: 40,
        id: "column operations on a table whose file states no widths give the default engine's package in the canonical session",
      },
    );
  });

  test("a suggested column insertion records the grid with its columns unmeasured, and rejecting restores no widths", async () => {
    const source = await packageWith("a table with no grid");
    const canonical = await viaCanonicalEngine(
      source,
      OPERATIONS["insertTableColumn after"]("20000001"),
      "tracked-changes",
    );
    expect(canonical.refusals).toEqual([]);
    expect(canonical.skipped).toEqual([]);
    const part = (await ZipArchive.read(canonical.bytes)).entry(
      "word/document.xml",
    );
    if (!part) throw new Error("The saved package has no document part.");
    const xml = await part.text();
    expect(xml).toMatch(
      /<w:tblGridChange w:id="\d+"><w:tblGrid><w:gridCol\/><w:gridCol\/><\/w:tblGrid><\/w:tblGridChange>/u,
    );
  });

  test("narrowing a column of spanning cells records the prior grid and cell spans", async () => {
    const source = await packageOf([
      paragraph("10000001", "Before"),
      table(
        [
          row(
            textCell(
              "20000001",
              "A1",
              '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>',
            ) + textCell("20000003", "B1"),
          ),
          row(
            textCell("20000002", "", '<w:gridSpan w:val="2"/><w:vMerge/>') +
              textCell("20000004", "B2"),
          ),
        ],
        3,
      ),
      paragraph("10000002", "After"),
    ]);
    const op = OPERATIONS.deleteTableColumn("20000001");
    const pending = await viaCanonicalEngine(source, op, "tracked-changes");
    expect(pending.refusals).toEqual([]);
    expect(pending.skipped).toEqual([]);
    const direct = await viaCanonicalEngine(source, op, "direct");
    expect(direct.skipped).toEqual([]);
    const part = (await ZipArchive.read(pending.bytes)).entry(
      "word/document.xml",
    );
    if (!part) throw new Error("The saved package has no document part.");
    const xml = await part.text();
    expect(xml.match(/<w:tcPrChange\b/gu)).toHaveLength(2);
    expect(xml).toMatch(/<w:tblGridChange\b/u);
    const reviewed = await reviewedInSession(pending.bytes);
    expect(reviewed.accepted).toEqual((await outcomes(direct.bytes)).accepted);
    expect(reviewed.rejected).toEqual((await outcomes(source)).rejected);
  });

  test("saved mixed column deletions preserve a neighboring vertical merge under review", async () => {
    const source = await packageOf([
      paragraph("10000001", "Before"),
      table(
        [
          row(
            textCell("20000001", "wide", '<w:gridSpan w:val="2"/>') +
              textCell("20000002", "merge", '<w:vMerge w:val="restart"/>'),
          ),
          row(
            textCell("20000003", "deleted") +
              textCell("20000004", "kept") +
              textCell("20000005", "", "<w:vMerge/>"),
          ),
        ],
        3,
      ),
      paragraph("10000002", "After"),
    ]);
    const op = OPERATIONS.deleteTableColumn("20000001");
    const pending = await viaCanonicalEngine(source, op, "tracked-changes");
    expect(pending.refusals).toEqual([]);
    expect(pending.skipped).toEqual([]);
    const direct = await viaCanonicalEngine(source, op, "direct");
    expect(direct.skipped).toEqual([]);
    const part = (await ZipArchive.read(pending.bytes)).entry(
      "word/document.xml",
    );
    if (!part) throw new Error("The saved package has no document part.");
    const xml = await part.text();
    expect(xml).toMatch(/<w:tcPrChange\b/u);
    expect(xml).toMatch(/<w:cellDel\b/u);
    expect(xml).toMatch(/<w:tblGridChange\b/u);
    const pendingDocument = await parseDocx(pending.bytes, {
      preloadFonts: false,
      detectVariables: false,
    });
    const pendingTable = createCanonicalSession(pendingDocument)
      .unwrap()
      .projection.doc.child(1);
    expect(pendingTable.firstChild?.lastChild?.attrs["rowspan"]).toBe(2);
    const reviewed = await reviewedInSession(pending.bytes);
    expect(reviewed.accepted).toEqual((await outcomes(direct.bytes)).accepted);
    expect(reviewed.rejected).toEqual((await outcomes(source)).rejected);
  });

  test("a saved unmeasured multi-column deletion has the right current and previous grids", async () => {
    const source = await packageOf([
      paragraph("10000001", "Before"),
      `<w:tbl>${TABLE_PROPERTIES}<w:tblGrid>${"<w:gridCol/>".repeat(3)}</w:tblGrid>${row(textCell("20000001", "A") + textCell("20000002", "B") + textCell("20000003", "C"))}</w:tbl>`,
      paragraph("10000002", "After"),
    ]);
    const loaded = await parseDocx(source, {
      preloadFonts: false,
      detectVariables: false,
    });
    const operation = {
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      story: OP_STORIES.MAIN,
      blockId: "20000003",
      column: 0,
      count: 2,
    } as const;
    const intent = { type: "table", operation } as const;
    const ids = createEditorIntentIdAllocator()(loaded, intent);
    const compiled = compileEditorIntent(loaded, {
      intent,
      mode: {
        type: "suggesting",
        revision: {
          id: ids.revisionId,
          author: "Reviewer",
          date: "2026-11-27T00:00:00Z",
        },
        newIds: ids.newIds,
      },
    }).unwrap();
    const pending = await createDocx(
      applyDocumentOps(loaded, compiled.ops).unwrap().document,
    );
    const part = (await ZipArchive.read(pending)).entry("word/document.xml");
    if (part === undefined)
      throw new TypeError("The saved package has no document part.");
    const xml = await part.text();
    expect(xml).toMatch(
      /<w:tblGrid><w:gridCol\/><w:tblGridChange\b[^>]*><w:tblGrid>(?:<w:gridCol\/>){3}<\/w:tblGrid>/u,
    );
    const direct = await createDocx(
      applyDocumentOp(loaded, operation).unwrap().document,
    );
    const reviewed = await reviewedInSession(pending);
    expect(reviewed.accepted).toEqual((await outcomes(direct)).accepted);
    expect(reviewed.rejected).toEqual((await outcomes(source)).rejected);
  });

  test.each([
    "merge",
    "merge with text",
    "merge absorbing a row",
    "merge retaining a row",
    "split",
  ] as const)(
    "a tracked horizontal %s survives save and review",
    async (kind) => {
      const merging = kind !== "split";
      const absorbsRow = kind === "merge absorbing a row";
      const retainsRow = kind === "merge retaining a row";
      let endBlockId = "20000002";
      if (absorbsRow) endBlockId = "20000006";
      if (retainsRow) endBlockId = "20000005";
      const source = await packageOf([
        paragraph("10000001", "Before"),
        table(
          [
            row(
              (merging
                ? textCell("20000001", "Owner") +
                  textCell(
                    "20000002",
                    kind === "merge with text" ? "Moved" : "",
                  )
                : textCell("20000001", "Owner", '<w:gridSpan w:val="2"/>')) +
                textCell("20000003", "Tail"),
            ),
            ...(absorbsRow || retainsRow
              ? [
                  row(
                    textCell("20000004", "Lower left") +
                      textCell("20000005", "Lower middle") +
                      textCell("20000006", "Lower right"),
                  ),
                ]
              : []),
          ],
          3,
        ),
        paragraph("10000002", "After"),
      ]);
      const op = merging
        ? ({
            type: "mergeTableCells",
            id: "merge",
            blockId: "20000001",
            endBlockId,
          } satisfies FolioDocumentOperation)
        : ({
            type: "splitTableCell",
            id: "split",
            blockId: "20000001",
          } satisfies FolioDocumentOperation);
      const pending = await viaCanonicalEngine(source, op, "tracked-changes");
      expect(pending.refusals).toEqual([]);
      expect(pending.skipped).toEqual([]);
      const direct = await viaCanonicalEngine(source, op, "direct");
      expect(direct.skipped).toEqual([]);
      const part = (await ZipArchive.read(pending.bytes)).entry(
        "word/document.xml",
      );
      if (!part) throw new Error("The saved package has no document part.");
      const xml = await part.text();
      expect(xml).toMatch(/<w:tcPrChange\b/u);
      expect(xml).toMatch(merging ? /<w:cellDel\b/u : /<w:cellIns\b/u);
      if (absorbsRow) expect(xml).toMatch(/<w:trPr><w:del\b/u);
      if (retainsRow) expect(xml).toMatch(/<w:cellIns\b/u);
      if (kind === "merge with text" || absorbsRow || retainsRow) {
        expect(xml).toMatch(/<w:moveFromRangeStart\b/u);
        expect(xml).toMatch(/<w:moveToRangeStart\b/u);
        expect(xml).toMatch(/<w:moveFrom\b/u);
        expect(xml).toMatch(/<w:moveTo\b/u);
      }
      const reviewed = await reviewedInSession(pending.bytes);
      expect(reviewed.accepted).toEqual(
        (await outcomes(direct.bytes)).accepted,
      );
      expect(reviewed.rejected).toEqual((await outcomes(source)).rejected);
    },
  );

  // Without the fix the canonical session refuses the suggestion the default
  // engine records; the parity check above then fails.
  test("the old refusal fails the parity check", async () => {
    const refused = async (): Promise<CanonicalRun> => {
      const source = await packageWith("a table with no grid");
      const run = await viaCanonicalEngine(
        source,
        OPERATIONS["insertTableColumn after"]("20000001"),
        "tracked-changes",
      );
      return {
        ...run,
        skipped: ["unsupportedBlock"],
        messages: ["Tracked column edits require an explicit previous grid."],
        refusals: [{ gap: CANONICAL_GAP.trackedTableStructure }],
      };
    };
    await expect(
      sameAsDefaultEngine(
        {
          body: "a table with no grid",
          anchor: "20000001",
          operation: "insertTableColumn after",
          mode: "tracked-changes",
        },
        refused,
      ),
    ).rejects.toThrow();
  });
});
