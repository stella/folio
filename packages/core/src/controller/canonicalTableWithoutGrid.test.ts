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

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { FolioDocumentOperation } from "../ai-edits";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import {
  type Mode,
  outcomes,
  packageOf,
  paragraph,
  reviewedInSession,
  row,
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
        refusals: [{ gap: CANONICAL_GAP.publicUnsupportedInline }],
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
