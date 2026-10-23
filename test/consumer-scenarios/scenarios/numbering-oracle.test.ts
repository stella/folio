import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import {
  hashFolioAIBlockText,
  normalizeFolioAIBlockText,
} from "@stll/folio-core/ai-edits/snapshot";
import type { FolioAIBlockParagraphProperties } from "@stll/folio-core/ai-edits/types";
import { paragraphNumberingFromSlots } from "@stll/folio-core/docx";
import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  createDocx,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  paragraph,
  parseDocx,
} from "@stll/folio-core/server";

import {
  decimalLevel,
  directNumberedDocument,
  emojiDocument,
  openReviewer,
  packDocument,
  storiesDocument,
  styleNumberedDocument,
  tableDocument,
  toArrayBuffer,
} from "../support/documents.ts";
import { runFlow } from "../support/fuzz.ts";
import {
  applyChecked,
  capture,
  compareWithModel,
  expectOperation,
  modelOf,
  numberingFactsOf,
  type Row,
  rowsOf,
} from "../support/oracle.ts";
import { modelPendingTextEffect } from "../support/pending-text-oracle.ts";

const slowTest = (name: string, fn: () => Promise<void>) =>
  test(name, { timeout: 30_000 }, fn);
const MARKER_VISIBILITY_TEST =
  "numbering kind follows marker visibility, independently of the retained reference";
const HIDDEN_LEVELS_TEST =
  "undefined and hidden numbering levels keep provenance without becoming list items";

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
      const expectedLevel =
        level?.kind === "reference" ? (level.ilvl ?? 0) : undefined;
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
                  listLevel:
                    expectedLevel === undefined ? 0 : expectedLevel + 1,
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
              row.id === "inserted"
                ? Object.assign({}, row, { listReference: undefined })
                : row,
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
  const listed = live.map((row) =>
    row === pending ? Object.assign({}, row, { text: "" }) : row,
  );
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
  const restyleKept = {
    type: "setBlockParagraphProperties",
    blockId: kept.id,
  } as const;
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    const model = modelOf(accepted, listed);
    model.numberingFacts = lacking;
    model.mode = mode;
    assert.throws(
      () =>
        expectOperation(model, {
          ...restyleKept,
          properties: { styleId: null },
        }),
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
        levels: [
          decimalLevel(0, "%1."),
          decimalLevel(1, "%1.%2."),
          decimalLevel(2, "%1.%2.%3."),
        ],
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

slowTest("markerless numbering regression seed 1339716102", async () => {
  await runFlow(1339716102, 20, "collisions", {
    generation: "targeted",
    swarm: "enabled",
  });
});

test("a pending new list does not make a same-text markerless paragraph visible", async () => {
  const repeatedText = "Repeated paragraph";
  const markerless = paragraph(repeatedText);
  markerless.formatting = {
    ...markerless.formatting,
    numPr: paragraphNumberingReference({ numId: 5, ilvl: 0 }),
  };
  const document = fromMarkdown("Anchor");
  document.package.document.content = [markerless, paragraph("Anchor")];
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        levels: [{ ...decimalLevel(0, ""), numFmt: "none" }],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  const reviewer = await openReviewer(await packDocument(document));
  const original = rowsOf(reviewer).find((row) => row.text === repeatedText);
  const anchor = rowsOf(reviewer).find((row) => row.text === "Anchor");
  assert.ok(original);
  assert.ok(anchor);
  assert.equal(original.kind, "paragraph");

  const staged = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: repeatedText,
        numbering: { kind: "newList", format: "bullet" },
      },
    ],
    "suggested",
    "stage a same-text list paragraph",
  );
  assert.equal(staged.applied.length, 1);
  const repeated = rowsOf(reviewer).filter((row) => row.text === repeatedText);
  const retained = repeated.find((row) => row.id === original.id);
  const inserted = repeated.find((row) => row.id !== original.id);
  assert.ok(retained);
  assert.ok(inserted);
  assert.equal(retained.kind, "paragraph");
  assert.equal(inserted.kind, "listItem");
  assert.ok(inserted.listReference);

  const captured = await capture(reviewer, "suggested");
  assert.equal(captured.numberingSource.type, "live");
  if (captured.numberingSource.type !== "live")
    throw new Error("Expected live provenance");
  assert.equal(
    captured.numberingSource.facts.markerLevels.get(5)?.has(0),
    false,
  );
  assert.equal(
    captured.numberingSource.facts.markerLevels
      .get(inserted.listReference.numId)
      ?.has(inserted.listReference.level),
    true,
  );
});

slowTest(MARKER_VISIBILITY_TEST, async () => {
  for (const marker of [
    "visible",
    "missing",
    "hidden",
    "visibleOverride",
    "hiddenOverride",
    "none",
    "noneLegal",
    "emptyDecimal",
  ] as const) {
    const level = marker === "missing" ? 8 : 0;
    const visible =
      marker === "visible" ||
      marker === "visibleOverride" ||
      marker === "noneLegal" ||
      marker === "emptyDecimal";
    const document = fromMarkdown("# Heading\n\nBody");
    document.package.numbering = {
      abstractNums: [
        {
          abstractNumId: 5,
          levels: [
            {
              ...decimalLevel(
                0,
                marker === "none" ||
                  marker === "noneLegal" ||
                  marker === "emptyDecimal"
                  ? ""
                  : "%1.",
              ),
              ...(marker === "none" ? { numFmt: "none" } : {}),
              ...(marker === "noneLegal"
                ? { numFmt: "none", isLgl: true }
                : {}),
              ...(marker === "hidden" || marker === "visibleOverride"
                ? { rPr: { hidden: true } }
                : {}),
            },
          ],
        },
      ],
      nums: [
        {
          numId: 5,
          abstractNumId: 5,
          ...(marker === "visibleOverride" || marker === "hiddenOverride"
            ? {
                levelOverrides: [
                  {
                    ilvl: 0,
                    lvl: {
                      ...decimalLevel(0, "%1."),
                      rPr: { hidden: marker === "hiddenOverride" },
                    },
                  },
                ],
              }
            : {}),
        },
      ],
    };
    assert.ok(document.package.styles);
    document.package.styles.styles.push({
      styleId: "MarkerList",
      type: "paragraph",
      pPr: { numPr: paragraphNumberingFromSlots({ numId: 5, ilvl: level }) },
    });
    for (const block of document.package.document.content) {
      if (block.type !== "paragraph") continue;
      block.formatting = { ...block.formatting, styleId: "MarkerList" };
    }
    const bytes = await packDocument(document);
    for (const anchorText of ["Heading", "Body"]) {
      for (const numbering of [
        paragraphNumberingReference({ numId: 5, ilvl: level }),
        { kind: "inherit" },
        { kind: "levelOnly", ilvl: level },
      ] as const) {
        for (const type of [
          "insertAfterBlock",
          "setBlockParagraphProperties",
        ] as const) {
          for (const mode of [
            "direct",
            "tracked-changes",
            "suggested",
          ] as const) {
            const reviewer = await openReviewer(bytes);
            const preRows = rowsOf(reviewer);
            const anchor = preRows.find(({ text }) => text === anchorText);
            assert.ok(anchor);
            const operation =
              type === "insertAfterBlock"
                ? { type, blockId: anchor.id, text: "Inserted", numbering }
                : { type, blockId: anchor.id, properties: { numbering } };
            const context = `${marker} ${anchorText} ${numbering.kind} ${type} ${mode}`;
            const model = modelOf(preRows);
            model.mode = mode;
            model.numberingFacts = await numberingFactsOf(bytes);
            expectOperation(model, operation);
            const result = await applyChecked(
              reviewer,
              [operation],
              mode,
              context,
            );
            assert.equal(result.applied.length, 1, context);
            assert.deepEqual(result.issues, [], context);
            const rows = rowsOf(reviewer);
            const targetText =
              type === "insertAfterBlock" ? "Inserted" : anchorText;
            const target = rows.find(({ text }) => text === targetText);
            assert.ok(target, context);
            const bodyKind = visible ? "listItem" : "paragraph";
            const expectedKind =
              anchorText === "Heading" ? "heading" : bodyKind;
            assert.equal(target.kind, expectedKind, context);
            assert.deepEqual(
              target.listReference,
              visible ? { numId: 5, level } : undefined,
              context,
            );
            assert.equal(
              target.listLevel,
              visible ? level : undefined,
              context,
            );
            assert.deepEqual(compareWithModel(model, rows), [], context);
            assert.match(
              compareWithModel(
                model,
                rows.map((row) =>
                  row === target
                    ? Object.assign({}, row, {
                        kind:
                          expectedKind === "paragraph"
                            ? "listItem"
                            : "paragraph",
                      })
                    : row,
                ),
              ).join("\n"),
              /kind/u,
              context,
            );
          }
        }
      }
    }
  }
});

test(
  "pending numbering provenance survives insertion and restyling across every editable story",
  { timeout: 60_000 },
  async () => {
    const document = await parseDocx(toArrayBuffer(await storiesDocument()), {
      preloadFonts: false,
    });
    document.package.numbering ??= { abstractNums: [], nums: [] };
    for (const numId of [900, 901]) {
      assert.ok(
        !document.package.numbering.nums.some((entry) => entry.numId === numId),
      );
      document.package.numbering.abstractNums.push({
        abstractNumId: numId,
        levels: [decimalLevel(0, "%1."), decimalLevel(1, "%1.%2.")],
      });
      document.package.numbering.nums.push({ numId, abstractNumId: numId });
    }
    assert.ok(document.package.styles);
    document.package.styles.styles.push(
      {
        styleId: "OracleList",
        type: "paragraph",
        pPr: { numPr: paragraphNumberingFromSlots({ numId: 900, ilvl: 0 }) },
      },
      {
        styleId: "OracleOtherList",
        type: "paragraph",
        pPr: { numPr: paragraphNumberingFromSlots({ numId: 901, ilvl: 0 }) },
      },
    );
    const bytes = await packDocument(document);
    const stories = (await openReviewer(bytes)).listStories();
    assert.deepEqual(
      new Set(stories.map(({ handle }) => handle.type)),
      new Set(["main", "header", "footer", "footnote", "endnote"]),
    );
    const cases = [
      {
        properties: {
          numbering: paragraphNumberingReference({ numId: 901, ilvl: 0 }),
        },
        direct: paragraphNumberingFromSlots({ numId: 901, ilvl: 0 }),
      },
      {
        properties: {
          numbering: paragraphNumberingReference({ numId: 900, ilvl: 1 }),
        },
        direct: paragraphNumberingFromSlots({ numId: 900, ilvl: 1 }),
      },
      {
        properties: {
          numbering: paragraphNumberingReference({ numId: 900, ilvl: 0 }),
        },
        direct: paragraphNumberingFromSlots({ numId: 900, ilvl: 0 }),
      },
      {
        properties: { numbering: { kind: "levelOnly", ilvl: 1 } },
        direct: { kind: "levelOnly", ilvl: 1 },
      },
      { properties: { numbering: { kind: "none" } }, direct: { kind: "none" } },
      { properties: { numbering: { kind: "inherit" } }, direct: undefined },
      { properties: { styleId: "Heading3" }, direct: undefined },
    ] as const satisfies readonly {
      properties: FolioAIBlockParagraphProperties;
      direct: ReturnType<typeof paragraphNumberingFromSlots>;
    }[];
    for (const { handle: story } of stories) {
      for (const { properties, direct } of cases) {
        for (const inheritFormatting of [true, false]) {
          const reviewer = await openReviewer(bytes);
          const anchor = rowsOf(reviewer, story).find(
            (row) => row.text.length > 0,
          );
          assert.ok(anchor);
          const initial = reviewer.applyDocumentOperationsToStory({
            undefinedReferences: "refuse",
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
            undefinedReferences: "refuse",
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
          if (pre.numberingSource.type !== "live")
            throw new Error("Expected live provenance");
          const facts = pre.numberingSource.facts;
          assert.equal(
            facts.styles.get("Heading2"),
            undefined,
            "Requested heading is unnumbered",
          );
          assert.deepEqual(
            facts.styles.get("OracleList"),
            paragraphNumberingFromSlots({ numId: 900, ilvl: 0 }),
          );
          for (const row of pre.rows.filter(
            (candidate) => candidate.kind !== "diagnostic",
          ))
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
            undefinedReferences: "refuse",
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
                  row.id === inserted.id
                    ? Object.assign({}, row, corrupted)
                    : row,
                ),
              ).join("\n"),
              /listLevel|numbering/u,
            );
          }
          const beforeStyle = await capture(reviewer, "suggested", { story });
          assert.equal(beforeStyle.numberingSource.type, "live");
          if (beforeStyle.numberingSource.type !== "live")
            throw new Error("Expected live provenance");
          const styleModel = modelOf(beforeStyle.rows);
          styleModel.numberingFacts = beforeStyle.numberingSource.facts;
          styleModel.mode = "suggested";
          const styleOperation = {
            id: "restyle",
            type: "setBlockParagraphProperties",
            blockId: anchor.id,
            properties: { styleId: "OracleOtherList" },
          } as const;
          expectOperation(styleModel, styleOperation);
          const restyled = reviewer.applyDocumentOperationsToStory({
            undefinedReferences: "refuse",
            story,
            batch: {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              operations: [styleOperation],
              mode: "suggested",
            },
          });
          assert.equal(restyled.applied.length, 1);
          assert.deepEqual(restyled.issues, []);
          assert.deepEqual(
            compareWithModel(styleModel, rowsOf(reviewer, story)),
            [],
          );
          const changed = rowsOf(reviewer, story).find(
            (row) => row.id === anchor.id,
          );
          assert.ok(changed);
          assert.deepEqual(
            changed.listReference,
            direct?.kind === "none"
              ? undefined
              : {
                  numId: direct?.kind === "reference" ? direct.numId : 901,
                  level: direct?.ilvl ?? 0,
                },
          );
          if (direct?.kind === "reference") {
            assert.match(
              compareWithModel(
                styleModel,
                rowsOf(reviewer, story).map((row) =>
                  row.id === anchor.id
                    ? Object.assign({}, row, {
                        listReference: {
                          numId: direct.numId === 900 ? 901 : 900,
                          level: direct.ilvl ?? 0,
                        },
                      })
                    : row,
                ),
              ).join("\n"),
              /numbering/u,
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
        numbering: { kind: "newList", format: "bullet" },
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

test("a reference resolves a list created by an earlier suggested insertion", async () => {
  const reviewer = await openReviewer(await tableDocument());
  const anchors = rowsOf(reviewer).filter((row) => row.text.length > 0);
  const firstAnchor = anchors.at(0);
  const secondAnchor = anchors.at(1);
  assert.ok(firstAnchor);
  assert.ok(secondAnchor);

  const created = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: firstAnchor.id,
        text: "Created bullet.",
        numbering: { kind: "newList", format: "bullet" },
      },
    ],
    "suggested",
    "create a pending list",
  );
  assert.equal(created.applied.length, 1);
  const bullet = rowsOf(reviewer).find((row) => row.text === "Created bullet.");
  assert.ok(bullet);
  assert.deepEqual(bullet.listReference, { numId: 1, level: 0 });

  const referenced = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: secondAnchor.id,
        text: "Referenced bullet.",
        numbering: { kind: "reference", numId: 1, ilvl: 0 },
      },
    ],
    "suggested",
    "reference a pending list",
  );
  assert.equal(referenced.applied.length, 1);
  const result = rowsOf(reviewer).find(
    (row) => row.text === "Referenced bullet.",
  );
  assert.ok(result);
  assert.equal(result.kind, "listItem");
  assert.equal(result.displayLabel, "•");
});

test("a reference resolves a pending list inserted beside a table cell", async () => {
  const reviewer = await openReviewer(await tableDocument());
  const tableCells = rowsOf(reviewer).filter((row) => row.table !== undefined);
  const listAnchor = tableCells.at(0);
  const referenceAnchor = tableCells.at(-1);
  assert.ok(listAnchor);
  assert.ok(referenceAnchor);

  await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: listAnchor.id,
        text: "Created beside a table.",
        numbering: { kind: "newList", format: "bullet" },
      },
    ],
    "suggested",
    "create a pending list beside a table cell",
  );
  const bullet = rowsOf(reviewer).find(
    (row) => row.text === "Created beside a table.",
  );
  assert.ok(bullet);
  assert.deepEqual(bullet.listReference, { numId: 1, level: 0 });

  await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: referenceAnchor.id,
        text: "Reference the pending list.",
        numbering: { kind: "reference", numId: 1, ilvl: 0 },
      },
    ],
    "suggested",
    "reference a pending list created beside a table cell",
  );
  const referenced = rowsOf(reviewer).find(
    (row) => row.text === "Reference the pending list.",
  );
  assert.ok(referenced);
  assert.equal(referenced.kind, "listItem");
  assert.equal(referenced.displayLabel, "•");
});

type PendingRecord = ReturnType<
  Awaited<ReturnType<typeof openReviewer>>["exportPendingSuggestions"]
>[number];

/** Text-only expectations for the pinned flow, independent of operation application. */
const expectedPendingLoads = (
  records: readonly PendingRecord[],
  savedText: ReadonlyMap<string, string>,
) => {
  const textByBlock = new Map(savedText);
  return records.map(({ suggestionId, anchor, operation }) => {
    const text = textByBlock.get(anchor.blockId);
    if (text === undefined)
      return {
        status: "stale",
        suggestionId,
        reason: "missingAnchor",
      } as const;
    if (
      hashFolioAIBlockText(normalizeFolioAIBlockText(text)) !==
      anchor.originalTextHash
    )
      return { status: "stale", suggestionId, reason: "textChanged" } as const;
    if (
      anchor.startOffset !== undefined &&
      anchor.endOffset !== undefined &&
      anchor.selectedTextHash !== undefined &&
      hashFolioAIBlockText(text.slice(anchor.startOffset, anchor.endOffset)) !==
        anchor.selectedTextHash
    )
      return { status: "stale", suggestionId, reason: "textChanged" } as const;
    modelPendingTextEffect(operation, textByBlock);
    return { status: "restaged", suggestionId } as const;
  });
};

test(
  "suggested collision numbering survives host persistence and tracked DOCX save",
  { timeout: 30_000 },
  async () => {
    let checked = false;
    await runFlow(1873083933, 16, "collisions", {
      generation: "targeted",
      checks: [
        {
          name: "numbered proposal persistence",
          check: async ({ flow, index, saved }) => {
            if (index !== 14) return;
            checked = true;
            const text = "Delivery payment clause.";
            const expected = { numId: 2, level: 0 };
            const inserted = rowsOf(flow.reviewer).find(
              (row) => row.text === text,
            );
            assert.ok(inserted);
            assert.deepEqual(inserted.listReference, expected);
            const records = flow.reviewer.exportPendingSuggestions();
            const proposal = records.find(
              ({ operation }) =>
                (operation.type === "insertAfterBlock" ||
                  operation.type === "insertBeforeBlock") &&
                operation.text === text,
            );
            assert.ok(proposal);
            const { bytes } = await saved();
            const reopened = await openReviewer(bytes);
            // Proposals are host-owned, outside DOCX until explicitly accepted.
            assert.ok(!rowsOf(reopened).some((row) => row.text === text));
            const savedText = new Map(
              reopened.snapshot().blocks.map((block) => [block.id, block.text]),
            );
            const expectedLoads = expectedPendingLoads(records, savedText);
            const loaded = reopened.loadPendingSuggestions(
              JSON.parse(JSON.stringify(records)),
            );
            assert.deepEqual(loaded, expectedLoads);
            // Each anchor's load state comes from the text oracle, including op-1
            // when its saved text is unchanged. The changed-text control below pins
            // the stale case without contradicting that source of truth.
            // Earlier proposals may change this anchor during replay. Stage the success
            // control against the saved baseline instead of requiring a stale record to load.
            // The live anchor inherits list numbering only through a pending deletion
            // that the saved baseline does not contain, so anchor the control on the
            // nearest saved paragraph that carries the same numbering.
            const { operation } = proposal;
            assert.ok(operation.type === "insertAfterBlock");
            const savedRows = rowsOf(await openReviewer(bytes));
            const liveAnchorAt = savedRows.findIndex(
              (row) => row.id === operation.blockId,
            );
            assert.ok(
              liveAnchorAt >= 0,
              "The live anchor exists in the saved baseline",
            );
            const numberedAnchor = savedRows
              .slice(0, liveAnchorAt + 1)
              .findLast((row) => row.listReference !== undefined);
            assert.ok(
              numberedAnchor,
              "The saved baseline keeps a numbered anchor",
            );
            assert.deepEqual(numberedAnchor.listReference, expected);
            const staging = await openReviewer(bytes);
            const staged = await applyChecked(
              staging,
              [
                {
                  ...operation,
                  id: "saved-heading-roundtrip",
                  blockId: numberedAnchor.id,
                },
              ],
              "suggested",
              "heading proposal against saved baseline",
            );
            assert.equal(staged.applied.length, 1);
            assert.deepEqual(
              rowsOf(staging).find((row) => row.text === text)?.listReference,
              expected,
              "Staging on the saved baseline inherits the anchor's numbering",
            );
            const freshRecords = staging.exportPendingSuggestions();
            assert.equal(freshRecords.length, 1);
            const freshProposal = freshRecords.at(0);
            assert.ok(freshProposal);
            const expectedFreshLoads = expectedPendingLoads(
              freshRecords,
              savedText,
            );
            assert.deepEqual(expectedFreshLoads, [
              { status: "restaged", suggestionId: freshProposal.suggestionId },
            ]);
            // Positive control: change the saved paragraph itself, not a load result.
            const changedDocument = await parseDocx(toArrayBuffer(bytes), {
              preloadFonts: false,
            });
            const changedParagraph =
              changedDocument.package.document.content.find(
                (block) =>
                  block.type === "paragraph" &&
                  block.paraId === freshProposal.anchor.paraId,
              );
            assert.ok(changedParagraph?.type === "paragraph");
            changedParagraph.content = [
              {
                type: "run",
                content: [{ type: "text", text: "Changed saved anchor." }],
              },
            ];
            const changed = await openReviewer(
              await packDocument(changedDocument),
            );
            const changedText = new Map(
              changed.snapshot().blocks.map((block) => [block.id, block.text]),
            );
            const staleControl = expectedPendingLoads(
              freshRecords,
              changedText,
            );
            assert.deepEqual(staleControl, [
              {
                status: "stale",
                suggestionId: freshProposal.suggestionId,
                reason: "textChanged",
              },
            ]);
            assert.deepEqual(
              changed.loadPendingSuggestions(freshRecords),
              staleControl,
            );
            const restaged = await openReviewer(bytes);
            assert.deepEqual(
              restaged.loadPendingSuggestions(
                JSON.parse(JSON.stringify(freshRecords)),
              ),
              expectedFreshLoads,
            );
            const restored = rowsOf(restaged).find((row) => row.text === text);
            assert.ok(restored);
            assert.deepEqual(restored.listReference, expected);
            assert.equal(
              restaged.acceptSuggestion(freshProposal.suggestionId),
              true,
            );
            const tracked = await openReviewer(
              new Uint8Array(await restaged.toBuffer()),
            );
            const persisted = rowsOf(tracked).find((row) => row.text === text);
            assert.ok(persisted);
            assert.deepEqual(persisted.listReference, expected);
            assert.ok(
              tracked.getChanges().some((change) => change.text.includes(text)),
            );
            tracked.acceptAll();
            const accepted = await openReviewer(
              new Uint8Array(await tracked.toBuffer()),
            );
            const authored = rowsOf(accepted).find((row) => row.text === text);
            assert.ok(authored);
            assert.deepEqual(authored.listReference, expected);
          },
        },
      ],
    });
    assert.equal(
      checked,
      true,
      "The generated seed must reach its reported insertion",
    );
  },
);

test("a pending inserted anchor has complete live numbering provenance", async () => {
  const reviewer = await openReviewer(await directNumberedDocument());
  const anchor = rowsOf(reviewer).find(
    (row) => row.text === "Unnumbered body text.",
  );
  assert.ok(anchor);
  const staged = await applyChecked(
    reviewer,
    [
      {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Pending numbered anchor",
        numbering: { kind: "reference", numId: 7, ilvl: 0 },
      },
    ],
    "suggested",
    "stage new numbered paragraph",
  );
  assert.equal(staged.applied.length, 1);
  const pending = rowsOf(reviewer).find(
    (row) => row.text === "Pending numbered anchor",
  );
  assert.ok(pending);
  const pre = await capture(reviewer, "suggested");
  assert.equal(pre.numberingSource.type, "live");
  if (pre.numberingSource.type !== "live")
    throw new Error("Expected live provenance");
  assert.deepEqual(
    pre.numberingSource.facts.direct.get(pending.id),
    paragraphNumberingFromSlots({ numId: 7, ilvl: 0 }),
  );
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

test("package numbering provenance covers allocated paragraph identities in every editable story", async () => {
  for (const identity of ["authored", "mixed", "positional"] as const) {
    const document = await parseDocx(toArrayBuffer(await storiesDocument()), {
      preloadFonts: false,
    });
    document.package.document.content.push(
      { ...paragraph("Repeated paragraph"), paraId: "71234501" },
      { ...paragraph("Repeated paragraph"), paraId: "71234502" },
    );
    const authoredBytes = new Uint8Array(await createDocx(document));
    const authoredReviewer = await openReviewer(authoredBytes);
    const authoredFacts = await numberingFactsOf(authoredBytes);
    let ordinal = 0;
    const removeIds = (value: unknown): void => {
      if (typeof value !== "object" || value === null) return;
      if ("paraId" in value) {
        if (
          identity === "positional" ||
          (identity === "mixed" && ordinal % 2 === 0)
        )
          delete value.paraId;
        ordinal++;
      }
      for (const child of Object.values(value)) removeIds(child);
      if (value instanceof Map)
        for (const child of value.values()) removeIds(child);
    };
    removeIds(document.package);
    const bytes = new Uint8Array(await createDocx(document));
    const reviewer = await openReviewer(bytes);
    const facts = await numberingFactsOf(bytes);
    const stories = reviewer.listStories();
    assert.deepEqual(
      new Set(stories.map(({ handle }) => handle.type)),
      new Set(["main", "header", "footer", "footnote", "endnote"]),
    );
    for (const { handle } of stories) {
      const rows = rowsOf(reviewer, handle).filter(
        (row) => row.kind !== "diagnostic",
      );
      const authoredRows = rowsOf(authoredReviewer, handle).filter(
        (row) => row.kind !== "diagnostic",
      );
      assert.deepEqual(
        rows.map(({ text }) => text),
        authoredRows.map(({ text }) => text),
      );
      assert.deepEqual(
        rows.map(({ id }) => facts.direct.get(id)),
        authoredRows.map(({ id }) => authoredFacts.direct.get(id)),
      );
      for (const row of rows) {
        assert.ok(
          facts.direct.has(row.id),
          `${identity} ${handle.type}: ${row.id}`,
        );
        const model = modelOf(rows);
        model.numberingFacts = facts;
        expectOperation(model, {
          type: "setBlockParagraphProperties",
          blockId: row.id,
          properties: { styleId: null },
        });
      }
    }
    const repeated = rowsOf(reviewer).filter(
      ({ text }) => text === "Repeated paragraph",
    );
    assert.equal(repeated.length, 2);
    assert.equal(new Set(repeated.map(({ id }) => id)).size, 2);
  }
});

slowTest(HIDDEN_LEVELS_TEST, async () => {
  for (const definitionSource of [
    "abstract",
    "override",
    "styleLink",
  ] as const) {
    const document = fromMarkdown(
      Array.from({ length: 9 }, (_, level) => `Anchor level ${level}.`).join(
        "\n\n",
      ),
    );
    // Exercise every legal level with visible, hidden and absent definitions.
    const levels = Array.from({ length: 9 }, (_, level) => level)
      .filter((level) => level % 3 !== 2)
      .map((level) =>
        Object.assign(decimalLevel(level, `%${level + 1}.`), {
          rPr: { hidden: level % 3 === 1 },
        }),
      );
    let baseLevels = levels;
    if (definitionSource === "override") {
      baseLevels = levels.map((level) =>
        Object.assign({}, level, { rPr: { hidden: !level.rPr.hidden } }),
      );
    }
    if (definitionSource === "styleLink") baseLevels = [];
    document.package.numbering = {
      abstractNums: [
        {
          abstractNumId: 7,
          levels: baseLevels,
          ...(definitionSource === "styleLink"
            ? { numStyleLink: "LinkedNumbering" }
            : {}),
        },
        ...(definitionSource === "styleLink"
          ? [{ abstractNumId: 8, styleLink: "LinkedNumbering", levels }]
          : []),
      ],
      nums: [
        {
          numId: 7,
          abstractNumId: 7,
          ...(definitionSource === "override"
            ? { levelOverrides: levels.map((lvl) => ({ ilvl: lvl.ilvl, lvl })) }
            : {}),
        },
      ],
    };
    for (let level = 0; level < 9; level++) {
      document.package.styles?.styles.push({
        styleId: `Level${level}`,
        type: "paragraph",
        pPr: { numPr: paragraphNumberingFromSlots({ numId: 7, ilvl: level }) },
      });
    }
    const bytes = await packDocument(document);
    const facts = await numberingFactsOf(bytes);
    for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
      for (const request of [
        "reference",
        "levelOnly",
        "inherit",
        "insert",
      ] as const) {
        const reviewer = await openReviewer(bytes);
        const anchors = rowsOf(reviewer);
        const operations = anchors.map((anchor, level) => {
          const reference = {
            kind: "reference",
            numId: 7,
            ilvl: level,
          } as const;
          switch (request) {
            case "insert":
              return {
                type: "insertAfterBlock",
                blockId: anchor.id,
                text: `Inserted level ${level}.`,
                numbering: reference,
              };
            case "reference":
              return {
                type: "setBlockParagraphProperties",
                blockId: anchor.id,
                properties: { numbering: reference },
              };
            case "levelOnly":
              return {
                type: "setBlockParagraphProperties",
                blockId: anchor.id,
                properties: {
                  styleId: `Level${level}`,
                  numbering: { kind: "levelOnly", ilvl: level },
                },
              };
            case "inherit":
              return {
                type: "setBlockParagraphProperties",
                blockId: anchor.id,
                properties: {
                  styleId: `Level${level}`,
                  numbering: { kind: "inherit" },
                },
              };
          }
        });
        const receipt = await applyChecked(
          reviewer,
          operations,
          mode,
          `${definitionSource}/${request}`,
        );
        assert.equal(receipt.applied.length, 9);
        if (mode !== "suggested") reviewer.acceptAll();
        const rows = rowsOf(reviewer);
        for (let level = 0; level < 9; level++) {
          const text = `${request === "insert" ? "Inserted" : "Anchor"} level ${level}.`;
          const row = rows.find((candidate) => candidate.text === text);
          assert.ok(row);
          assert.equal(row.kind, level % 3 === 0 ? "listItem" : "paragraph");
          assert.deepEqual(
            row.listReference,
            level % 3 === 0 ? { numId: 7, level } : undefined,
          );
          const model = modelOf(anchors);
          model.numberingFacts = facts;
          // Sensitivity: the oracle must reject both a spurious list and lost list classification.
          const operation = operations.at(level);
          assert.ok(operation);
          expectOperation(model, operation);
          const expectedRows = anchors.map((anchor) =>
            anchor.id === row.id ? row : anchor,
          );
          if (request === "insert") expectedRows.splice(level + 1, 0, row);
          assert.deepEqual(compareWithModel(model, expectedRows), []);
          assert.match(
            compareWithModel(
              model,
              expectedRows.map((candidate) =>
                candidate === row
                  ? Object.assign({}, candidate, {
                      kind: row.kind === "listItem" ? "paragraph" : "listItem",
                    })
                  : candidate,
              ),
            ).join("\n"),
            /kind/u,
          );
        }
      }
    }
  }
});

for (const seed of [1339716047, 1339716277]) {
  test(`missing numbering level regression seed ${seed} (#1819)`, async () => {
    await runFlow(seed, 20, "random", {
      generation: "targeted",
      swarm: "enabled",
    });
  });
}
