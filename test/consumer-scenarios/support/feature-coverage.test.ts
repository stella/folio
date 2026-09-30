import assert from "node:assert/strict";
import { test } from "node:test";

import { createRandom } from "./random.ts";
import {
  addFeatureHit,
  emptyFeatureCells,
  emptyFeatureCoverage,
  gapWeight,
  generatedSelection,
  operationSelection,
  placementSelection,
  shapeFeatureSignature,
  targetFeatureSignature,
  weightedChoice,
} from "./feature-coverage.ts";

test("feature signature distinguishes structure at the touched target", () => {
  assert.deepEqual(
    targetFeatureSignature(
      {
        kind: "listItem",
        listLevel: 2,
        displayLabel: "3.2",
        table: { outerTableIndex: 0, tableIndex: 1, columnSpan: 2, rowSpan: 1 },
      },
      new Set(["field", "commentAnchor", "pendingRevision"]),
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
  assert.deepEqual(shapeFeatureSignature(["footnote", "endnote"]), ["footnote"]);
  assert.deepEqual(shapeFeatureSignature(["list-decimal", "list-bullet"]), ["none"]);
});

test("selection signatures retain whole-document and cell selections", () => {
  assert.equal(placementSelection("document"), "whole-document");
  assert.equal(placementSelection("cross-paragraph", true), "cell-selection");
  assert.equal(operationSelection({ range: { blockId: "A1B2C3D4" } }), "paragraph-range");
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
  assert.ok(emptyFeatureCells(coverage).includes("replaceRange | field | whole-document"));
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
