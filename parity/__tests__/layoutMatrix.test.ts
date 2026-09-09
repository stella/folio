import { describe, expect, test } from "bun:test";

import { buildLayoutInteractionMatrix } from "../fixtures/layout-interaction-matrix";
import { summarizeLayoutMatrix } from "../layoutMatrix";
import type { CorpusReport, Divergence, FeatureAttributedResult } from "../types";

const scenarios = buildLayoutInteractionMatrix().slice(0, 2);

const resultFor = (
  id: string,
  divergences: Divergence[] = [],
  rasterStatus: "match" | "difference" | "dimension-mismatch" = "match",
): FeatureAttributedResult => ({
  file: `/synthetic/${id}.docx`,
  score: divergences.length === 0 ? 1 : 0.5,
  referencePages: 1,
  folioPages: 1,
  totalReferenceLines: 1,
  matchedLines: divergences.length === 0 ? 1 : 0,
  medianYOffsetPt: 0,
  divergences,
  attributed: [],
  docFeatures: [],
  fontEnvironment: { status: "native", tags: [], comparedLines: 1, matchingLines: 1 },
  rasterComparison: {
    status: "compared",
    score: rasterStatus === "match" ? 1 : 0.9,
    diffPixels: rasterStatus === "match" ? 0 : 10,
    totalPixels: 100,
    pages: [
      rasterStatus === "dimension-mismatch"
        ? {
            status: "dimension-mismatch",
            page: 1,
            referenceWidthPx: 100,
            referenceHeightPx: 100,
            folioWidthPx: 90,
            folioHeightPx: 100,
            diffPixels: 10,
            totalPixels: 100,
            similarity: 0.9,
          }
        : {
            status: rasterStatus,
            page: 1,
            widthPx: 100,
            heightPx: 100,
            diffPixels: rasterStatus === "match" ? 0 : 10,
            totalPixels: 100,
            similarity: rasterStatus === "match" ? 1 : 0.9,
          },
    ],
  },
});

const reportFor = (
  results: FeatureAttributedResult[],
  reference: "word" | "libreoffice" = "word",
) =>
  ({
    generatedAt: "2026-09-09T00:00:00.000Z",
    reference: { id: reference, displayName: reference },
    results,
    clusters: [],
  }) satisfies CorpusReport;

describe("reference layout matrix summary", () => {
  test("separates required structural failures from advisory raster differences", () => {
    const first = scenarios.at(0)!;
    const second = scenarios.at(1)!;
    const report = summarizeLayoutMatrix(
      reportFor([
        resultFor(first.id, [
          { kind: "pagination", text: "Synthetic", referencePage: 1, folioPage: 2 },
        ]),
        resultFor(second.id, [], "difference"),
      ]),
      scenarios,
    );

    expect(report.summary).toEqual({ total: 2, pass: 0, advisory: 1, fail: 1 });
    expect(report.cases.at(0)?.requiredDivergences).toEqual({ pagination: 1 });
    expect(report.requiredInteractionClusters.length).toBeGreaterThan(0);
    expect(
      report.requiredInteractionClusters.every(({ caseIds }) => caseIds.includes(first.id)),
    ).toBeTrue();
  });

  test("makes page-raster shape failures required", () => {
    const first = scenarios.at(0)!;
    const report = summarizeLayoutMatrix(
      reportFor([resultFor(first.id, [], "dimension-mismatch")]),
      [first],
    );

    expect(report.summary.fail).toBe(1);
    expect(report.cases.at(0)?.requiredRasterFailures).toBe(1);
  });

  test("does not gate line flow when font geometry is unverified", () => {
    const first = scenarios.at(0)!;
    const result = resultFor(first.id, [
      { kind: "line-break", page: 1, referenceTexts: ["A"], folioTexts: ["A", "B"] },
    ]);
    result.fontEnvironment = { status: "mismatch", tags: [], comparedLines: 1, matchingLines: 0 };
    const report = summarizeLayoutMatrix(reportFor([result]), [first]);

    expect(report.summary).toEqual({ total: 1, pass: 0, advisory: 1, fail: 0 });
  });

  test("supports each reference and rejects incomplete case sets", () => {
    const first = scenarios.at(0)!;
    expect(
      summarizeLayoutMatrix(reportFor([resultFor(first.id)], "libreoffice"), [first]).reference.id,
    ).toBe("libreoffice");
    expect(() => summarizeLayoutMatrix(reportFor([]), [first])).toThrow(
      `Missing matrix result: ${first.id}`,
    );
  });
});
