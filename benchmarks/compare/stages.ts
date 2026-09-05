/**
 * The four stages of one comparison, timed separately.
 *
 * The stage functions are the ones {@link compareDocx} itself composes, not a
 * copy of them: a second pipeline written for the benchmark would drift from
 * the real one, and the numbers would then describe code nobody ships.
 */

import {
  applyComparison,
  parseComparison,
  planComparison,
  serializeComparison,
} from "@stll/folio-core/compare/compare";
import type {
  CompareChange,
  CompareDocxOptions,
  CompareUnsupportedPart,
} from "@stll/folio-core/compare/types";

export const COMPARE_STAGES = Object.freeze(["parse", "align", "apply", "serialize"] as const);

export type CompareStage = (typeof COMPARE_STAGES)[number];

export type StageDurations = Record<CompareStage, number>;

export type StagedCompareRun = {
  status: "ok";
  durations: StageDurations;
  buffer: ArrayBuffer;
  changes: readonly CompareChange[];
  /** Package parts the comparison did not look at, by reason. */
  unsupported: readonly CompareUnsupportedPart[];
  /** Blocks the base story carried, the size the timings should be read against. */
  baseBlocks: number;
};

export type StagedCompareFailure = {
  status: "failed";
  stage: CompareStage;
  error: string;
};

export type StagedCompareResult = StagedCompareRun | StagedCompareFailure;

const describe = (error: unknown): string =>
  error instanceof Error ? `${error.name}: ${error.message}` : String(error);

/** Run one comparison, recording where each millisecond went. */
export const runStagedCompare = async (
  base: ArrayBuffer,
  target: ArrayBuffer,
  options: CompareDocxOptions,
): Promise<StagedCompareResult> => {
  const parseStart = performance.now();
  const parsed = await parseComparison(base, target, options);
  const parse = performance.now() - parseStart;
  if (parsed.isErr()) {
    return { status: "failed", stage: "parse", error: describe(parsed.error) };
  }

  const alignStart = performance.now();
  const planned = planComparison(parsed.value);
  const align = performance.now() - alignStart;
  if (planned.isErr()) {
    return { status: "failed", stage: "align", error: describe(planned.error) };
  }

  const applyStart = performance.now();
  const changes = applyComparison(parsed.value, planned.value);
  const apply = performance.now() - applyStart;
  if (changes.isErr()) {
    return { status: "failed", stage: "apply", error: describe(changes.error) };
  }

  const serializeStart = performance.now();
  const serialized = await serializeComparison(parsed.value);
  const serialize = performance.now() - serializeStart;
  if (serialized.isErr()) {
    return { status: "failed", stage: "serialize", error: describe(serialized.error) };
  }

  return {
    status: "ok",
    durations: { parse, align, apply, serialize },
    buffer: serialized.value,
    changes: changes.value,
    unsupported: parsed.value.unsupported,
    baseBlocks: parsed.value.pairs.reduce(
      (total, { baseSnapshot }) => total + baseSnapshot.blocks.length,
      0,
    ),
  };
};
