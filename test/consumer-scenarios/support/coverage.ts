/**
 * The coverage ledger: which operation types met which stories, modes,
 * target features and session states. Every operation the oracle checks is
 * recorded (support/oracle.ts); a scenario process writes its ledger to
 * FOLIO_SCENARIO_COVERAGE_DIR when it exits, and the runner merges them,
 * prints a table, writes the summary and fails a required cell with no hits
 * (coverage-expectations.json). Imports nothing from folio: the runner reads
 * it from the monorepo.
 */

import { mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import {
  addFeatureHit,
  emptyFeatureCoverage,
  mergeFeatureCoverage,
  type FeatureCell,
  type FeatureCoverage,
} from "./feature-coverage.ts";

export const DIMENSIONS = ["op", "story", "mode", "feature", "step"] as const;
export type Dimension = (typeof DIMENSIONS)[number];

/** One cell: an operation type in a story, in a mode, at a feature, in a session state. */
export type Cell = Record<Dimension, string>;

/** How often a cell was applied (and checked), and refused. */
export type Count = { applied: number; refused: number };

export type Ledger = { version: 1; cells: Record<string, Count> };

/** The session a step runs in: the reviewer the fixture opened, a reopened one, or a new author's. */
export type StepKind = "fresh" | "reopened" | "newReviewer";

const SEPARATOR = " | ";
const keyOf = (cell: Cell): string =>
  DIMENSIONS.map((dimension) => cell[dimension]).join(SEPARATOR);
export const cellOf = (key: string): Cell => {
  const parts = key.split(SEPARATOR);
  return Object.fromEntries(
    DIMENSIONS.map((dimension, index) => [dimension, parts[index] ?? ""]),
  ) as Cell;
};

const ledger: Ledger = { version: 1, cells: {} };
const featureCoverage = emptyFeatureCoverage();

export const recordFeatureHit = (cell: FeatureCell): void => addFeatureHit(featureCoverage, cell);
export const registerFeatureOperations = (operations: readonly string[]): void => {
  for (const operation of operations) {
    if (!featureCoverage.operations.includes(operation)) featureCoverage.operations.push(operation);
  }
};

type HitObserver = (key: string, applied: boolean) => void;
const observers = new Set<HitObserver>();

export const recordHit = (cell: Cell, applied: boolean): void => {
  const key = keyOf(cell);
  const count = (ledger.cells[key] ??= { applied: 0, refused: 0 });
  if (applied) count.applied += 1;
  else count.refused += 1;
  for (const observer of observers) observer(key, applied);
};

/** Call `observer` on every hit from now on (a flow's signature); returns the unsubscribe. */
export const observeHits = (observer: HitObserver): (() => void) => {
  observers.add(observer);
  return () => {
    observers.delete(observer);
  };
};

const outputDir = process.env["FOLIO_SCENARIO_COVERAGE_DIR"];
if (outputDir) {
  process.on("exit", () => {
    if (Object.keys(ledger.cells).length === 0) return;
    mkdirSync(outputDir, { recursive: true });
    writeFileSync(
      path.join(outputDir, `ledger-${process.pid}-${Date.now()}.json`),
      JSON.stringify(ledger),
    );
    writeFileSync(
      path.join(outputDir, `features-${process.pid}-${Date.now()}.json`),
      JSON.stringify(featureCoverage),
    );
  });
}

// ---------------------------------------------------------------------------
// Merging, checking and printing (the runner)
// ---------------------------------------------------------------------------

/** Every ledger a scenario process wrote into `dir`, merged. */
export const mergeLedgers = (dir: string): Ledger => {
  const merged: Ledger = { version: 1, cells: {} };
  let files: string[] = [];
  try {
    files = readdirSync(dir).filter((file) => file.startsWith("ledger-") && file.endsWith(".json"));
  } catch {
    return merged;
  }
  for (const file of files) {
    const part = JSON.parse(readFileSync(path.join(dir, file), "utf8")) as Ledger;
    for (const [key, count] of Object.entries(part.cells)) {
      const into = (merged.cells[key] ??= { applied: 0, refused: 0 });
      into.applied += count.applied;
      into.refused += count.refused;
    }
  }
  return merged;
};

export const mergeFeatureFiles = (dir: string): FeatureCoverage => {
  let files: string[];
  try {
    files = readdirSync(dir).filter(
      (file) => file.startsWith("features-") && file.endsWith(".json"),
    );
  } catch {
    return emptyFeatureCoverage();
  }
  return mergeFeatureCoverage(
    files.map((file) => JSON.parse(readFileSync(path.join(dir, file), "utf8")) as FeatureCoverage),
  );
};

/** A cell pattern: a dimension left out, or `"*"`, matches anything. */
export type Pattern = Partial<Cell>;

export type Expectations = {
  /** Cells the default CI run must apply at least once. */
  required: { cell: Pattern; why?: string }[];
  /** Cells the public API cannot reach yet, and why. */
  unreachable: { cell: Pattern; reason: string }[];
  /** Limits of the ledger itself: what it cannot tell apart, and why. */
  notes?: string[];
};

const matches = (pattern: Pattern, cell: Cell): boolean =>
  DIMENSIONS.every((dimension) => {
    const wanted = pattern[dimension];
    return wanted === undefined || wanted === "*" || wanted === cell[dimension];
  });

export const describePattern = (pattern: Pattern): string =>
  DIMENSIONS.map((dimension) => `${dimension}=${pattern[dimension] ?? "*"}`).join(" ");

/** The total count of every cell `pattern` matches. */
export const countOf = (merged: Ledger, pattern: Pattern): Count => {
  const total = { applied: 0, refused: 0 };
  for (const [key, count] of Object.entries(merged.cells)) {
    if (!matches(pattern, cellOf(key))) continue;
    total.applied += count.applied;
    total.refused += count.refused;
  }
  return total;
};

/** Required cells nothing applied, each named. */
export const missingCells = (merged: Ledger, expectations: Expectations): string[] =>
  expectations.required
    .filter(({ cell }) => countOf(merged, cell).applied === 0)
    .map(({ cell, why }) => `${describePattern(cell)}${why ? ` (${why})` : ""}`);

/** Declared unreachable patterns with any attempted operation, applied or refused. */
export const hitUnreachableCells = (merged: Ledger, expectations: Expectations) =>
  expectations.unreachable
    .map(({ cell, reason }) => ({ cell, reason, ...countOf(merged, cell) }))
    .filter(({ applied, refused }) => applied + refused > 0);

/** Applied counts summed over every dimension but `rows` and `columns`. */
const pivot = (merged: Ledger, rows: Dimension, columns: Dimension) => {
  const table = new Map<string, Map<string, number>>();
  const columnKeys = new Set<string>();
  for (const [key, count] of Object.entries(merged.cells)) {
    const cell = cellOf(key);
    columnKeys.add(cell[columns]);
    const row = table.get(cell[rows]) ?? new Map<string, number>();
    row.set(cell[columns], (row.get(cell[columns]) ?? 0) + count.applied);
    table.set(cell[rows], row);
  }
  return { table, columnKeys: [...columnKeys].sort() };
};

const formatPivot = (merged: Ledger, rows: Dimension, columns: Dimension): string => {
  const { table, columnKeys } = pivot(merged, rows, columns);
  const rowKeys = [...table.keys()].sort();
  const width = Math.max(rows.length, ...rowKeys.map((key) => key.length));
  const widths = columnKeys.map((key) => Math.max(key.length, 5));
  const line = (label: string, values: string[]) =>
    `  ${label.padEnd(width)}  ${values.map((value, index) => value.padStart(widths[index] ?? 5)).join("  ")}`;
  return [
    `applied operations, ${rows} × ${columns}:`,
    line(rows, columnKeys),
    ...rowKeys.map((key) =>
      line(
        key,
        columnKeys.map((column) => String(table.get(key)?.get(column) ?? 0).replace(/^0$/u, "·")),
      ),
    ),
  ].join("\n");
};

/** The table the runner prints. */
export const formatLedger = (merged: Ledger): string => {
  const totals = countOf(merged, {});
  return [
    `coverage ledger: ${Object.keys(merged.cells).length} cells, ${totals.applied} applied, ${totals.refused} refused`,
    formatPivot(merged, "op", "story"),
    formatPivot(merged, "feature", "mode"),
    formatPivot(merged, "story", "step"),
  ].join("\n\n");
};

/** The JSON summary the runner writes: the merged cells and the marginals. */
export const summarize = (merged: Ledger, expectations: Expectations) => ({
  ...merged,
  marginals: Object.fromEntries(
    DIMENSIONS.map((dimension) => {
      const values: Record<string, Count> = {};
      for (const [key, count] of Object.entries(merged.cells)) {
        const value = cellOf(key)[dimension];
        const into = (values[value] ??= { applied: 0, refused: 0 });
        into.applied += count.applied;
        into.refused += count.refused;
      }
      return [dimension, values];
    }),
  ),
  required: expectations.required.map(({ cell }) => ({
    cell: describePattern(cell),
    ...countOf(merged, cell),
  })),
  unreachable: expectations.unreachable.map(({ cell, reason }) => ({
    cell: describePattern(cell),
    reason,
    ...countOf(merged, cell),
  })),
});
