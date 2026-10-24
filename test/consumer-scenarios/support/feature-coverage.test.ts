import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import { paragraph, run, table } from "@stll/folio-core/server";

import { openReviewer, packDocument } from "./documents.ts";
import {
  addFeatureHit,
  emptyFeatureCells,
  emptyFeatureCoverage,
  featureCoverageWeight,
  gapWeight,
  generatedSelection,
  operationSelection,
  placementSelection,
  shapeFeatureSignature,
  targetFeatureSignature,
  weightedChoice,
} from "./feature-coverage.ts";
import { createRandom } from "./random.ts";
import { featureIndex, storyKindOf } from "./targets.ts";
import type { Feature } from "./targets.ts";

test("feature signature distinguishes structure at the touched target", () => {
  const featureSignature = targetFeatureSignature(
    undefined,
    new Set<Feature>(["field", "commentAnchor", "pendingRevision"]),
    "body",
  );
  assert.deepEqual(
    targetFeatureSignature(
      undefined,
      new Set<Feature>(["pendingRevision", "commentAnchor", "field"]),
      "body",
    ),
    featureSignature,
  );
  assert.deepEqual(featureSignature, [
    "field",
    "comment-anchor",
    "tracked-change",
  ]);
  assert.deepEqual(
    targetFeatureSignature(
      {
        kind: "listItem",
        listReference: { numId: 1, level: 2 },
        displayLabel: "3.2",
        table: { outerTableIndex: 0, tableIndex: 1, columnSpan: 2, rowSpan: 1 },
      },
      new Set<Feature>(["field", "commentAnchor", "pendingRevision"]),
      "body",
    ),
    [
      "plain-table",
      "merged-cells",
      "nested-table",
      "numbered-list",
      "multi-level-list",
      "field",
      "comment-anchor",
      "tracked-change",
    ],
  );
  assert.deepEqual(shapeFeatureSignature(["footnote", "endnote"]), [
    "footnote",
  ]);
  assert.deepEqual(shapeFeatureSignature(["list-decimal", "list-bullet"]), [
    "none",
  ]);
  assert.deepEqual(
    targetFeatureSignature(
      undefined,
      new Set<Feature>(["noteReference"]),
      "footnote",
    ),
    ["footnote"],
  );
  assert.deepEqual(
    targetFeatureSignature(
      undefined,
      new Set<Feature>(["noteReference"]),
      "endnote",
    ),
    ["endnote"],
  );
});

test("body note references are a distinct target feature", async () => {
  const document = fromMarkdown("Ordinary body text.");
  const noteId = 37;
  document.package.footnotes = [
    { type: "footnote", id: noteId, content: [paragraph("Footnote.")] },
  ];
  const endnoteId = 38;
  document.package.endnotes = [
    { type: "endnote", id: endnoteId, content: [paragraph("Endnote.")] },
  ];
  const noteTable = table({ rows: [["Referenced table cell"]] });
  const noteCell = noteTable.rows.at(0)?.cells.at(0);
  assert.ok(noteCell);
  noteCell.content = [
    paragraph([
      run("Referenced table cell "),
      { type: "run", content: [{ type: "endnoteRef", id: endnoteId }] },
    ]),
  ];
  document.package.document.content.push(
    paragraph([
      run("Referenced body text "),
      { type: "run", content: [{ type: "footnoteRef", id: noteId }] },
    ]),
    noteTable,
    paragraph("Trailing body text."),
  );

  const reviewer = await openReviewer(await packDocument(document));
  const index = featureIndex(reviewer);
  const referenceBlocks = reviewer
    .getContent()
    .filter((block) => index.features.get(block.id)?.has("noteReference"));
  assert.equal(referenceBlocks.length, 2);
  assert.deepEqual(
    new Set(referenceBlocks.map((block) => storyKindOf(index, block))),
    new Set(["body", "tableCell"]),
  );
  for (const referenceBlock of referenceBlocks) {
    assert.ok(
      targetFeatureSignature(
        referenceBlock,
        index.features.get(referenceBlock.id) ?? new Set(),
        storyKindOf(index, referenceBlock),
      ).includes("note-reference"),
    );
  }

  const ordinaryBlock = reviewer
    .getContent()
    .find((block) => block.text === "Ordinary body text.");
  assert.ok(ordinaryBlock);
  const ordinaryFeatures = targetFeatureSignature(
    ordinaryBlock,
    index.features.get(ordinaryBlock.id) ?? new Set(),
    storyKindOf(index, ordinaryBlock),
  );
  assert.deepEqual(ordinaryFeatures, ["none"]);
});

test("selection signatures retain whole-document and cell selections", () => {
  assert.equal(placementSelection("document"), "whole-document");
  assert.equal(placementSelection("node"), "node-selection");
  assert.equal(placementSelection("cross-paragraph", true), "cell-selection");
  assert.equal(
    operationSelection({ range: { blockId: "A1B2C3D4" } }),
    "paragraph-range",
  );
  assert.equal(generatedSelection("splitBlock"), "caret");
});

test("rare cells receive bounded deterministic steering and empty cells remain visible", () => {
  const coverage = emptyFeatureCoverage();
  addFeatureHit(coverage, {
    operation: "replaceRange",
    feature: "field",
    selection: "paragraph-range",
  });
  assert.equal(coverage.cells["replaceRange | field | paragraph-range"], 1);
  assert.ok(
    emptyFeatureCells(coverage).includes(
      "replaceRange | field | whole-document",
    ),
  );
  assert.equal(gapWeight(0), 9);
  assert.equal(gapWeight(100), 1);
  const draw = (seed: number) => {
    const random = createRandom(seed);
    return Array.from({ length: 100 }, () =>
      weightedChoice(["empty", "common"], random, (candidate) =>
        gapWeight(candidate === "empty" ? 0 : 100),
      ),
    );
  };
  assert.deepEqual(draw(107), draw(107));
  assert.ok(draw(107).filter((choice) => choice === "empty").length > 70);
});

test("coverage weights steer each operation toward under-hit body note references", () => {
  const coverage = emptyFeatureCoverage();
  for (const operation of ["replaceRange", "formatRange"]) {
    for (let hit = 0; hit < 100; hit += 1) {
      addFeatureHit(coverage, {
        operation,
        feature: "none",
        selection: "paragraph-range",
      });
    }
  }
  for (let hit = 0; hit < 8; hit += 1) {
    addFeatureHit(coverage, {
      operation: "replaceRange",
      feature: "note-reference",
      selection: "paragraph-range",
    });
  }

  const noteReference = {
    kind: "paragraph",
    headingLevel: undefined,
    listReference: undefined,
    displayLabel: undefined,
  };
  const weight = (operation: string, known: ReadonlySet<Feature>) =>
    featureCoverageWeight({
      operation,
      target: noteReference,
      known,
      story: "body",
      selection: "paragraph-range",
      coverage,
    });

  assert.equal(weight("replaceRange", new Set<Feature>(["noteReference"])), 1);
  assert.equal(weight("formatRange", new Set<Feature>(["noteReference"])), 9);
  assert.equal(weight("replaceRange", new Set<Feature>()), 1);
  assert.equal(weight("formatRange", new Set<Feature>()), 1);
});
