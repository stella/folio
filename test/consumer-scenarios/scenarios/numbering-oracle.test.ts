import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";

import {
  decimalLevel,
  directNumberedDocument,
  openReviewer,
  packDocument,
  styleNumberedDocument,
} from "../support/documents.ts";
import {
  compareWithModel,
  expectOperation,
  modelOf,
  numberingFactsOf,
  rowsOf,
  type Row,
} from "../support/oracle.ts";

test("style numbering comes from definitions, preserves direct overrides, and catches missing numbering", async () => {
  for (const fixture of [directNumberedDocument, styleNumberedDocument]) {
    const bytes = await fixture();
    const rows = rowsOf(await openReviewer(bytes));
    const facts = await numberingFactsOf(bytes);
    const anchor = rows.find((row) => row.listReference !== undefined);
    assert.ok(anchor);
    for (const inheritFormatting of [undefined, true, false]) {
      const model = modelOf(rows);
      model.numberingFacts = facts;
      expectOperation(model, {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Inserted clause",
        styleId: "Heading2",
        ...(inheritFormatting === undefined ? {} : { inheritFormatting }),
      });
      const level =
        inheritFormatting === false
          ? facts.styles.get("Heading2")
          : (facts.direct.get(anchor.id) ?? facts.styles.get("Heading2"));
      const expectedLevel = level?.kind === "reference" ? (level.ilvl ?? 0) : undefined;
      const inserted: Row = {
        ...anchor,
        id: "inserted",
        text: "Inserted clause",
        styleId: "Heading2",
        kind: "heading",
        headingLevel: 2,
        listLevel: expectedLevel,
        previewRuns: [{ text: "Inserted clause", bold: true }],
        listReference:
          level?.kind === "reference"
            ? { numId: level.numId, level: expectedLevel ?? 0 }
            : undefined,
      };
      const result = rows.slice();
      result.splice(rows.indexOf(anchor) + 1, 0, inserted);
      assert.deepEqual(compareWithModel(model, result), []);
      assert.match(
        compareWithModel(
          model,
          result.map((row) =>
            row.id === "inserted"
              ? Object.assign({}, row, {
                  listLevel: expectedLevel === undefined ? 0 : expectedLevel + 1,
                })
              : row,
          ),
        ).join("\n"),
        /listLevel/u,
      );
      if (level?.kind === "reference") {
        assert.match(
          compareWithModel(
            model,
            result.map((row) =>
              row.id === "inserted" ? Object.assign({}, row, { listReference: undefined }) : row,
            ),
          ).join("\n"),
          /numbering/u,
        );
      }
    }
  }
});

test("style inheritance folds independent numbering slots and cancellation", async () => {
  const document = fromMarkdown("# Title\n\nBody");
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        levels: [decimalLevel(0, "%1."), decimalLevel(1, "%1.%2."), decimalLevel(2, "%1.%2.%3.")],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  document.package.styles?.styles.push(
    {
      styleId: "BaseList",
      type: "paragraph",
      pPr: { numPr: paragraphNumberingFromSlots({ numId: 5, ilvl: 1 }) },
    },
    {
      styleId: "NestedList",
      type: "paragraph",
      basedOn: "BaseList",
      pPr: { numPr: paragraphNumberingFromSlots({ ilvl: 2 }) },
    },
    {
      styleId: "CancelledList",
      type: "paragraph",
      basedOn: "NestedList",
      pPr: { numPr: paragraphNumberingFromSlots({ numId: 0 }) },
    },
  );
  const facts = await numberingFactsOf(await packDocument(document));
  assert.equal(facts.styles.get("NestedList")?.kind, "reference");
  assert.deepEqual(
    facts.styles.get("NestedList"),
    paragraphNumberingFromSlots({ numId: 5, ilvl: 2 }),
  );
  assert.deepEqual(facts.styles.get("CancelledList"), { kind: "none" });
});
