import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, parseDocx } from "@stll/folio-core/server";
import type { FolioAIBlockParagraphProperties } from "@stll/folio-core/ai-edits/types";

import { runFlow } from "../support/fuzz.ts";

import {
  decimalLevel,
  directNumberedDocument,
  openReviewer,
  packDocument,
  storiesDocument,
  emojiDocument,
  toArrayBuffer,
  styleNumberedDocument,
} from "../support/documents.ts";
import {
  applyChecked,
  capture,
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
  const listed = live.map((row) => (row === pending ? Object.assign({}, row, { text: "" }) : row));
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
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    const model = modelOf(accepted, listed);
    model.numberingFacts = lacking;
    model.mode = mode;
    assert.throws(
      () => expectOperation(model, { ...restyleKept, properties: { styleId: null } }),
      /Numbering provenance missing/u,
    );
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

test(
  "pending numbering provenance stays strict across every editable story",
  { timeout: 60_000 },
  async () => {
    const document = await parseDocx(toArrayBuffer(await storiesDocument()), {
      preloadFonts: false,
    });
    document.package.numbering ??= { abstractNums: [], nums: [] };
    for (const numId of [900, 901]) {
      assert.ok(!document.package.numbering.nums.some((entry) => entry.numId === numId));
      document.package.numbering.abstractNums.push({
        abstractNumId: numId,
        levels: [decimalLevel(0, "%1."), decimalLevel(1, "%1.%2.")],
      });
      document.package.numbering.nums.push({ numId, abstractNumId: numId });
    }
    assert.ok(document.package.styles);
    document.package.styles.styles.push({
      styleId: "OracleList",
      type: "paragraph",
      pPr: { numPr: { kind: "reference", numId: 900, ilvl: 0 } },
    });
    const bytes = await packDocument(document);
    const stories = (await openReviewer(bytes)).listStories();
    assert.deepEqual(
      new Set(stories.map(({ handle }) => handle.type)),
      new Set(["main", "header", "footer", "footnote", "endnote"]),
    );
    const cases = [
      {
        properties: { numbering: { numId: 901, level: 0 } },
        direct: { kind: "reference", numId: 901, ilvl: 0 },
      },
      {
        properties: { numbering: { numId: 900, level: 1 } },
        direct: { kind: "reference", numId: 900, ilvl: 1 },
      },
      { properties: { numbering: null }, direct: { kind: "none" } },
      { properties: { styleId: "Heading3" }, direct: undefined },
    ] as const satisfies readonly {
      properties: FolioAIBlockParagraphProperties;
      direct: ReturnType<typeof paragraphNumberingFromSlots>;
    }[];
    for (const { handle: story } of stories) {
      for (const { properties, direct } of cases) {
        for (const inheritFormatting of [true, false]) {
          const reviewer = await openReviewer(bytes);
          const anchor = rowsOf(reviewer, story).find((row) => row.text.length > 0);
          assert.ok(anchor);
          const initial = reviewer.applyDocumentOperationsToStory({
            story,
            batch: {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              operations: [
                {
                  id: "style",
                  type: "setBlockParagraphProperties",
                  blockId: anchor.id,
                  properties: { styleId: "OracleList" },
                },
              ],
              mode: "direct",
            },
          });
          assert.equal(initial.applied.length, 1);
          const pending = reviewer.applyDocumentOperationsToStory({
            story,
            batch: {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              operations: [
                {
                  id: "pending",
                  type: "setBlockParagraphProperties",
                  blockId: anchor.id,
                  properties,
                },
              ],
              mode: "suggested",
            },
          });
          assert.equal(pending.applied.length, 1);
          const pre = await capture(reviewer, "suggested", { story });
          assert.equal(pre.numberingSource.type, "live");
          if (pre.numberingSource.type !== "live") throw new Error("Expected live provenance");
          const facts = pre.numberingSource.facts;
          for (const row of pre.rows.filter((candidate) => candidate.kind !== "diagnostic"))
            assert.ok(facts.direct.has(row.id));
          assert.deepEqual(facts.direct.get(anchor.id), direct);
          const model = modelOf(pre.rows);
          model.numberingFacts = facts;
          model.mode = "suggested";
          const operation = {
            id: "insert",
            type: "insertAfterBlock",
            blockId: anchor.id,
            text: "Inserted requested heading",
            styleId: "Heading2",
            inheritFormatting,
          } as const;
          expectOperation(model, operation);
          const result = reviewer.applyDocumentOperationsToStory({
            story,
            batch: {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              operations: [operation],
              mode: "suggested",
            },
          });
          assert.equal(result.applied.length, 1);
          assert.deepEqual(result.issues, []);
          const actual = rowsOf(reviewer, story);
          assert.deepEqual(compareWithModel(model, actual), []);
          const inserted = actual.find((row) => row.text === operation.text);
          assert.ok(inserted);
          for (const corrupted of [
            { listLevel: (inserted.listLevel ?? 0) + 1 },
            {
              listReference: {
                numId: inserted.listReference?.numId === 900 ? 901 : 900,
                level: inserted.listReference?.level ?? 0,
              },
            },
          ]) {
            assert.match(
              compareWithModel(
                model,
                actual.map((row) =>
                  row.id === inserted.id ? Object.assign({}, row, corrupted) : row,
                ),
              ).join("\n"),
              /listLevel|numbering/u,
            );
          }
        }
      }
    }
  },
);

test("direct heading insertion beside a newly authored bullet keeps its direct numbering", async () => {
  const reviewer = await openReviewer(await emojiDocument());
  const anchor = rowsOf(reviewer).find((row) => row.text.length > 0);
  assert.ok(anchor);
  const first = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "A second-level heading.",
        styleId: "Heading2",
      },
      {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "A new bullet.",
        numbering: { start: "new", kind: "bullet" },
      },
    ],
    "direct",
    "legacy heading and bullet batch",
  );
  assert.equal(first.applied.length, 2);
  const bullet = rowsOf(reviewer).find((row) => row.text === "A new bullet.");
  assert.ok(bullet);
  const second = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: bullet.id,
        text: "Period buyer clause supplier supplier payment written.",
        styleId: "Heading2",
      },
    ],
    "direct",
    "legacy direct-numbered heading insertion",
  );
  assert.equal(second.applied.length, 1);
});

test("suggested collision regression preserves the anchor's pending direct numbering", async () => {
  await runFlow(1873083933, 16, "collisions", { generation: "targeted" });
});

test("a pending inserted anchor has complete live numbering provenance", async () => {
  const reviewer = await openReviewer(await directNumberedDocument());
  const anchor = rowsOf(reviewer).find((row) => row.text === "Unnumbered body text.");
  assert.ok(anchor);
  const staged = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Pending numbered anchor",
        numbering: { numId: 7, level: 0 },
      },
    ],
    "suggested",
    "stage new numbered paragraph",
  );
  assert.equal(staged.applied.length, 1);
  const pending = rowsOf(reviewer).find((row) => row.text === "Pending numbered anchor");
  assert.ok(pending);
  const pre = await capture(reviewer, "suggested");
  assert.equal(pre.numberingSource.type, "live");
  if (pre.numberingSource.type !== "live") throw new Error("Expected live provenance");
  assert.deepEqual(pre.numberingSource.facts.direct.get(pending.id), {
    kind: "reference",
    numId: 7,
    ilvl: 0,
  });
  const inserted = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: pending.id,
        text: "Heading after proposal",
        styleId: "Heading2",
      },
    ],
    "suggested",
    "inherit proposal numbering",
  );
  assert.equal(inserted.applied.length, 1);
});
