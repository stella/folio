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

test("a paragraph pending deletion has no numbering provenance and none is expected", async () => {
  const bytes = await directNumberedDocument();
  const live = rowsOf(await openReviewer(bytes));
  const pending = live.find((row) => row.listReference !== undefined);
  assert.ok(pending);
  // The accepted pre-state no longer holds the paragraph; the reviewer lists it blank.
  const accepted = live.filter((row) => row !== pending);
  const listed = live.map((row) => (row === pending ? { ...row, text: "" } : row));
  const facts = await numberingFactsOf(bytes);
  facts.direct.delete(pending.id);

  for (const styleId of [null, "Heading2"]) {
    const model = modelOf(accepted, listed);
    model.numberingFacts = facts;
    expectOperation(model, {
      type: "setBlockParagraphProperties",
      blockId: pending.id,
      properties: { styleId },
    });
    assert.deepEqual(model.unmodelled, []);
    assert.deepEqual(compareWithModel(model, accepted), []);
  }

  // A paragraph the pre-state holds must still have its provenance.
  const kept = accepted.find((row) => row.text.length > 0);
  assert.ok(kept);
  const lacking = await numberingFactsOf(bytes);
  lacking.direct.delete(kept.id);
  const restyleKept = { type: "setBlockParagraphProperties", blockId: kept.id } as const;
  for (const mode of ["direct", "tracked-changes"] as const) {
    const model = modelOf(accepted, listed);
    model.numberingFacts = lacking;
    model.mode = mode;
    assert.throws(
      () => expectOperation(model, { ...restyleKept, properties: { styleId: null } }),
      /Numbering provenance missing/u,
    );
  }

  // Suggested mode reads the document without its suggestions: a paragraph a
  // pending suggestion added is not in it, and its numbering is left open.
  const suggested = modelOf(live);
  suggested.numberingFacts = lacking;
  suggested.mode = "suggested";
  expectOperation(suggested, { ...restyleKept, properties: { styleId: "Heading2" } });
  assert.deepEqual(suggested.unmodelled, []);
  const restyled = (listLevel: number | undefined, listReference: Row["listReference"]): Row[] =>
    live.map((row) =>
      row.id === kept.id
        ? {
            ...row,
            styleId: "Heading2",
            kind: "heading",
            headingLevel: 2,
            listLevel,
            listReference,
          }
        : row,
    );
  for (const result of [restyled(undefined, undefined), restyled(1, { numId: 5, level: 1 })]) {
    assert.deepEqual(
      compareWithModel(suggested, result).filter((problem) => /listLevel|numbering/u.test(problem)),
      [],
    );
  }
  // The rest of the request is still held to.
  assert.match(
    compareWithModel(
      suggested,
      live.map((row) =>
        row.id === kept.id ? Object.assign({}, row, { styleId: "Heading3" }) : row,
      ),
    ).join("\n"),
    /styleId/u,
  );
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
