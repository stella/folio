import path from "node:path";

import { TaggedError } from "better-result";
import {
  layoutInteractionCasePairs,
  type LayoutInteractionCase,
} from "./fixtures/layout-interaction-matrix";
import type {
  CorpusReport,
  Divergence,
  DivergenceKind,
  FeatureAttributedResult,
  RasterComparison,
  ReferenceRendererInfo,
} from "./types";
import { isGeometryScoreReliable } from "./types";

export const LAYOUT_MATRIX_REPORT_SCHEMA = "folio.reference-layout-matrix";
export const LAYOUT_MATRIX_REPORT_VERSION = 1;

const REQUIRED_DIVERGENCE_KINDS = new Set<DivergenceKind>([
  "page-count",
  "pagination",
  "line-break",
  "missing-line",
  "extra-line",
  "text-mismatch",
]);

export type LayoutMatrixCaseStatus = "pass" | "advisory" | "fail";

type DivergenceCounts = Partial<Record<DivergenceKind, number>>;

export type LayoutMatrixCaseResult = {
  scenario: LayoutInteractionCase;
  status: LayoutMatrixCaseStatus;
  geometryReliable: boolean;
  referencePages: number;
  folioPages: number;
  requiredDivergences: DivergenceCounts;
  diagnosticDivergences: DivergenceCounts;
  requiredRasterFailures: number;
  rasterComparison?: RasterComparison;
};

export type LayoutMatrixInteractionCluster = {
  pair: string;
  failingCases: number;
  caseIds: string[];
};

export type ReferenceLayoutMatrixReport = {
  schema: typeof LAYOUT_MATRIX_REPORT_SCHEMA;
  version: typeof LAYOUT_MATRIX_REPORT_VERSION;
  generatedAt: string;
  reference: ReferenceRendererInfo;
  summary: Record<LayoutMatrixCaseStatus | "total", number>;
  cases: LayoutMatrixCaseResult[];
  requiredInteractionClusters: LayoutMatrixInteractionCluster[];
};

export class LayoutMatrixError extends TaggedError("LayoutMatrixError")<{
  message: string;
}> {}

export const layoutMatrixError = (message: string): LayoutMatrixError =>
  new LayoutMatrixError({ message });

const countDivergences = (divergences: Divergence[]): DivergenceCounts => {
  const counts: DivergenceCounts = {};
  for (const { kind } of divergences) counts[kind] = (counts[kind] ?? 0) + 1;
  return counts;
};

const hasRasterDifference = (raster: RasterComparison | undefined): boolean =>
  raster === undefined || raster.status === "empty" || raster.score < 1;

const countRequiredRasterFailures = (raster: RasterComparison | undefined): number => {
  if (raster === undefined || raster.status === "empty") return 1;
  return raster.pages.filter(
    ({ status }) =>
      status === "dimension-mismatch" ||
      status === "missing-reference" ||
      status === "missing-folio",
  ).length;
};

const caseStatus = (
  hasRequiredFailure: boolean,
  hasDiagnostic: boolean,
): LayoutMatrixCaseStatus => {
  if (hasRequiredFailure) return "fail";
  if (hasDiagnostic) return "advisory";
  return "pass";
};

const summarizeCase = (
  scenario: LayoutInteractionCase,
  result: FeatureAttributedResult,
): LayoutMatrixCaseResult => {
  const geometryReliable = isGeometryScoreReliable(result.fontEnvironment);
  const required = result.divergences.filter(({ kind }) => REQUIRED_DIVERGENCE_KINDS.has(kind));
  const diagnostics = result.divergences.filter(({ kind }) => !REQUIRED_DIVERGENCE_KINDS.has(kind));
  const requiredRasterFailures = countRequiredRasterFailures(result.rasterComparison);
  const hasRequiredFailure =
    (geometryReliable && required.length > 0) || requiredRasterFailures > 0;
  const hasDiagnostic =
    required.length > 0 || diagnostics.length > 0 || hasRasterDifference(result.rasterComparison);
  const status = caseStatus(hasRequiredFailure, hasDiagnostic);

  return {
    scenario,
    status,
    geometryReliable,
    referencePages: result.referencePages,
    folioPages: result.folioPages,
    requiredDivergences: countDivergences(required),
    diagnosticDivergences: countDivergences(diagnostics),
    requiredRasterFailures,
    ...(result.rasterComparison === undefined ? {} : { rasterComparison: result.rasterComparison }),
  };
};

const resultCaseId = ({ file }: FeatureAttributedResult): string =>
  path.basename(file, path.extname(file));

const requireCaseResults = (
  scenarios: readonly LayoutInteractionCase[],
  results: FeatureAttributedResult[],
): Map<string, FeatureAttributedResult> => {
  const expected = new Set(scenarios.map(({ id }) => id));
  const byId = new Map<string, FeatureAttributedResult>();
  for (const result of results) {
    const id = resultCaseId(result);
    if (!expected.has(id)) throw layoutMatrixError(`Unexpected matrix result: ${id}.`);
    if (byId.has(id)) throw layoutMatrixError(`Duplicate matrix result: ${id}.`);
    byId.set(id, result);
  }
  for (const { id } of scenarios) {
    if (!byId.has(id)) throw layoutMatrixError(`Missing matrix result: ${id}.`);
  }
  return byId;
};

const clusterFailures = (cases: LayoutMatrixCaseResult[]): LayoutMatrixInteractionCluster[] => {
  const clusters = new Map<string, string[]>();
  for (const result of cases) {
    if (result.status !== "fail") continue;
    for (const pair of layoutInteractionCasePairs(result.scenario)) {
      const ids = clusters.get(pair) ?? [];
      ids.push(result.scenario.id);
      clusters.set(pair, ids);
    }
  }
  return Array.from(clusters, ([pair, caseIds]) => ({
    pair,
    failingCases: caseIds.length,
    caseIds,
  })).toSorted(
    (left, right) => right.failingCases - left.failingCases || left.pair.localeCompare(right.pair),
  );
};

export const summarizeLayoutMatrix = (
  report: CorpusReport,
  scenarios: readonly LayoutInteractionCase[],
): ReferenceLayoutMatrixReport => {
  const byId = requireCaseResults(scenarios, report.results);
  const cases = scenarios.map((scenario) => {
    const result = byId.get(scenario.id);
    if (!result) throw layoutMatrixError(`Missing matrix result: ${scenario.id}.`);
    return summarizeCase(scenario, result);
  });
  const summary = { total: cases.length, pass: 0, advisory: 0, fail: 0 };
  for (const result of cases) summary[result.status] += 1;

  return {
    schema: LAYOUT_MATRIX_REPORT_SCHEMA,
    version: LAYOUT_MATRIX_REPORT_VERSION,
    generatedAt: report.generatedAt,
    reference: report.reference,
    summary,
    cases,
    requiredInteractionClusters: clusterFailures(cases),
  };
};
