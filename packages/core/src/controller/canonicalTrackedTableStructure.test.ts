/**
 * Table, row and column operations on tables holding more than words, a
 * nested table, bookmarks, content controls and an earlier suggestion, give
 * in the canonical session the package the default engine gives: pending,
 * every change accepted, every change rejected, and resolved in the
 * canonical session itself.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";

import { propertyTestTimeout } from "../../../../test/property-testing";
import type { FolioDocumentOperation } from "../ai-edits";
import {
  cell,
  deletedParagraph,
  deletedRow,
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

setDefaultTimeout(propertyTestTimeout(600_000));

const BEFORE = paragraph("10000001", "Before");
const AFTER = paragraph("10000002", "After");

/** Bodies holding one table between two paragraphs; the cell paragraphs are the anchors. */
const BODIES = {
  "a nested table": {
    xml: [
      BEFORE,
      table(
        [
          row(
            cell(
              paragraph("20000001", "A1") +
                table(
                  [
                    row(
                      textCell("21000001", "N1") + textCell("21000002", "N2"),
                    ),
                  ],
                  2,
                ) +
                paragraph("20000002", ""),
            ) + textCell("20000003", "B1"),
          ),
          row(textCell("20000004", "A2") + textCell("20000005", "B2")),
        ],
        2,
      ),
      AFTER,
    ],
    anchors: ["20000001", "21000001", "20000003", "20000004"],
  },
  "bookmarks in cells": {
    xml: [
      BEFORE,
      table(
        [
          row(
            cell(
              '<w:p w14:paraId="30000001"><w:bookmarkStart w:id="1" w:name="inline"/><w:r><w:t>A1</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
            ) +
              cell(
                '<w:bookmarkStart w:id="2" w:name="block"/><w:bookmarkEnd w:id="2"/>' +
                  paragraph("30000002", "B1"),
              ),
          ),
          row(textCell("30000003", "A2") + textCell("30000004", "B2")),
        ],
        2,
      ),
      AFTER,
    ],
    anchors: ["30000001", "30000002", "30000003"],
  },
  "content controls in a row": {
    xml: [
      BEFORE,
      table(
        [
          row(
            `<w:sdt><w:sdtPr><w:id w:val="11"/></w:sdtPr><w:sdtContent>${textCell("40000001", "A1")}</w:sdtContent></w:sdt>` +
              cell(
                `<w:sdt><w:sdtPr><w:id w:val="12"/></w:sdtPr><w:sdtContent>${paragraph("40000002", "B1")}</w:sdtContent></w:sdt>` +
                  paragraph("40000003", ""),
              ),
          ),
          row(textCell("40000004", "A2") + textCell("40000005", "B2")),
        ],
        2,
      ),
      AFTER,
    ],
    anchors: ["40000001", "40000002", "40000004"],
  },
  "a row already suggested for deletion": {
    xml: [
      BEFORE,
      table(
        [
          row(textCell("50000001", "A1") + textCell("50000002", "B1")),
          deletedRow(
            cell(deletedParagraph("50000003", "A2", 21)) +
              cell(deletedParagraph("50000004", "B2", 23)),
            20,
          ),
          row(textCell("50000005", "A3") + textCell("50000006", "B3")),
        ],
        2,
      ),
      AFTER,
    ],
    anchors: ["50000001", "50000002", "50000005"],
  },
} as const;

type Body = keyof typeof BODIES;

const OPERATIONS = {
  deleteTable: (blockId: string) => ({
    id: "op",
    type: "deleteTable",
    blockId,
  }),
  deleteTableRow: (blockId: string) => ({
    id: "op",
    type: "deleteTableRow",
    blockId,
  }),
  deleteTableColumn: (blockId: string) => ({
    id: "op",
    type: "deleteTableColumn",
    blockId,
  }),
  "insertTableRow after": (blockId: string) => ({
    id: "op",
    type: "insertTableRow",
    blockId,
    position: "after",
    cellTexts: ["New"],
  }),
  "insertTableColumn before": (blockId: string) => ({
    id: "op",
    type: "insertTableColumn",
    blockId,
    position: "before",
  }),
} as const satisfies Record<
  string,
  (blockId: string) => FolioDocumentOperation
>;

const MODES: readonly Mode[] = ["direct", "tracked-changes"];

describe("table operations on tables holding more than words", () => {
  test("give the default engine's package in the canonical session, pending, accepted, rejected and resolved there", async () => {
    const differences: string[] = [];
    for (const body of Object.keys(BODIES) as Body[]) {
      const source = await packageOf(BODIES[body].xml);
      for (const anchor of BODIES[body].anchors)
        for (const operation of Object.keys(
          OPERATIONS,
        ) as (keyof typeof OPERATIONS)[])
          for (const mode of MODES) {
            const label = `${body} / ${anchor} / ${operation} / ${mode}`;
            try {
              await compare(label, source, OPERATIONS[operation](anchor), mode);
            } catch (error) {
              differences.push(
                `${label}: threw ${error instanceof Error ? error.message : String(error)}`,
              );
            }
          }
    }
    expect(differences).toEqual([]);

    async function compare(
      label: string,
      source: ArrayBuffer,
      op: FolioDocumentOperation,
      mode: Mode,
    ): Promise<void> {
      const legacy = await viaDefaultEngine(source, op, mode);
      const canonical = await viaCanonicalEngine(source, op, mode);
      const refusals = canonical.refusals.map((refusal) =>
        JSON.stringify(refusal),
      );
      if (refusals.length > 0) {
        differences.push(`${label}: refused ${refusals.join("; ")}`);
        return;
      }
      if (
        JSON.stringify(canonical.skipped) !== JSON.stringify(legacy.skipped)
      ) {
        differences.push(
          `${label}: skipped ${JSON.stringify(canonical.skipped)}, the default engine ${JSON.stringify(legacy.skipped)}`,
        );
        return;
      }
      if (legacy.skipped.length > 0) return;
      const expected = await outcomes(legacy.bytes);
      const actual = await outcomes(canonical.bytes);
      for (const key of ["pending", "accepted", "rejected"] as const)
        if (JSON.stringify(actual[key]) !== JSON.stringify(expected[key]))
          differences.push(
            `${label}: ${key} ${JSON.stringify(actual[key])}, the default engine ${JSON.stringify(expected[key])}`,
          );
      if (mode === "direct") return;
      const session = await reviewedInSession(canonical.bytes);
      for (const key of ["accepted", "rejected"] as const)
        if (JSON.stringify(session[key]) !== JSON.stringify(expected[key]))
          differences.push(
            `${label}: ${key} in the session ${JSON.stringify(session[key])}, the default engine ${JSON.stringify(expected[key])}`,
          );
    }
  });
});
