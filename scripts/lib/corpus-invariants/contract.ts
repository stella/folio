/**
 * The contract every extended corpus invariant implements.
 *
 * The first four invariants the gate shipped with all run off one parse. So do
 * these, and for the same reason: a 4.5 MB package takes tens of seconds to
 * parse, so an invariant that re-parsed the input would cost more than it
 * proves. Each one receives the already-parsed model and the original bytes,
 * and returns failures plus the wall time of every stage it ran.
 *
 * Budgets are advisory, not preemptive. Nothing here can interrupt a
 * synchronous serializer mid-call, so a stage that overruns is recorded as a
 * failure after the fact and the file's remaining stages are skipped; the
 * worker deadline in `corpus-pool.ts` remains the only hard stop. Recording the
 * overrun is the point: a stage that needs a minute on one package is a
 * performance cliff, which is a finding of its own.
 */

import type { Document } from "@stll/folio-core/types/document";

import type { CorpusFailure, CorpusInvariant } from "../corpus-signature";

export const EXTENDED_CORPUS_INVARIANTS = {
  reserialize: "reserialize",
  editorRoundTrip: "editor-round-trip",
  editLocality: "edit-locality",
  saveIdempotence: "save-idempotence",
  schemaValidity: "schema-validity",
  pipelineTotality: "pipeline-totality",
  kernelDifferential: "kernel-differential",
  performance: "performance",
} as const;

export type ExtendedCorpusInvariant =
  (typeof EXTENDED_CORPUS_INVARIANTS)[keyof typeof EXTENDED_CORPUS_INVARIANTS];

/**
 * Which baseline file owns a signature.
 *
 * One family per file, so that a change to one invariant's findings never
 * rewrites another's, and a branch that adds an invariant never conflicts with
 * a branch that re-measures an existing one.
 */
export const CORPUS_INVARIANT_FAMILIES = {
  /** The gate's original four plus `completes`, which keep `corpus/baseline.json`. */
  core: "core",
  reserialize: "reserialize",
  editorRoundTrip: "editor-round-trip",
  editLocality: "edit-locality",
  saveIdempotence: "save-idempotence",
  schemaValidity: "schema-validity",
  pipelineTotality: "pipeline-totality",
  kernelDifferential: "kernel-differential",
  performance: "performance",
} as const;

export type CorpusInvariantFamily =
  (typeof CORPUS_INVARIANT_FAMILIES)[keyof typeof CORPUS_INVARIANT_FAMILIES];

/**
 * Whether a family's findings ratchet, or only report.
 *
 * The ratchet is exact in both directions: a signature that gains files is a
 * regression, and one that loses them must be written down before the gate
 * passes again. That rule needs a verdict that depends on the input alone. A
 * wall-clock verdict does not: the same file on a loaded machine crosses a
 * budget it clears on an idle one, so a baseline measured under load fails a
 * quiet run and a baseline measured idle fails a busy one. Recording timings
 * and ratcheting them are different jobs, and only the first is sound here.
 *
 * A report-only family is measured, written to the census and printed in the
 * report with its outliers. It owns no baseline file and is never compared.
 * Deterministic performance guards are separate work.
 */
export const CORPUS_FAMILY_GATING = {
  [CORPUS_INVARIANT_FAMILIES.core]: "gating",
  [CORPUS_INVARIANT_FAMILIES.reserialize]: "gating",
  [CORPUS_INVARIANT_FAMILIES.editorRoundTrip]: "gating",
  [CORPUS_INVARIANT_FAMILIES.editLocality]: "gating",
  [CORPUS_INVARIANT_FAMILIES.saveIdempotence]: "gating",
  [CORPUS_INVARIANT_FAMILIES.schemaValidity]: "gating",
  [CORPUS_INVARIANT_FAMILIES.pipelineTotality]: "gating",
  [CORPUS_INVARIANT_FAMILIES.kernelDifferential]: "gating",
  [CORPUS_INVARIANT_FAMILIES.performance]: "report-only",
} as const satisfies Record<CorpusInvariantFamily, "gating" | "report-only">;

export type CorpusFamilyGating = (typeof CORPUS_FAMILY_GATING)[CorpusInvariantFamily];

export const isGatingFamily = (family: CorpusInvariantFamily): boolean =>
  CORPUS_FAMILY_GATING[family] === "gating";

export const EXTENDED_INVARIANT_FAMILY = {
  [EXTENDED_CORPUS_INVARIANTS.reserialize]: CORPUS_INVARIANT_FAMILIES.reserialize,
  [EXTENDED_CORPUS_INVARIANTS.editorRoundTrip]: CORPUS_INVARIANT_FAMILIES.editorRoundTrip,
  [EXTENDED_CORPUS_INVARIANTS.editLocality]: CORPUS_INVARIANT_FAMILIES.editLocality,
  [EXTENDED_CORPUS_INVARIANTS.saveIdempotence]: CORPUS_INVARIANT_FAMILIES.saveIdempotence,
  [EXTENDED_CORPUS_INVARIANTS.schemaValidity]: CORPUS_INVARIANT_FAMILIES.schemaValidity,
  [EXTENDED_CORPUS_INVARIANTS.pipelineTotality]: CORPUS_INVARIANT_FAMILIES.pipelineTotality,
  [EXTENDED_CORPUS_INVARIANTS.kernelDifferential]: CORPUS_INVARIANT_FAMILIES.kernelDifferential,
  [EXTENDED_CORPUS_INVARIANTS.performance]: CORPUS_INVARIANT_FAMILIES.performance,
} as const satisfies Record<ExtendedCorpusInvariant, CorpusInvariantFamily>;

/**
 * Which baseline file owns a signature. It lives beside the map it reads so
 * that both censuses can classify a failure without importing each other.
 */
export const familyOf = (invariant: CorpusInvariant): CorpusInvariantFamily =>
  // SAFETY: the map is total over the extended invariants; anything else is a
  // core invariant, which keeps `corpus/baseline.json`.
  EXTENDED_INVARIANT_FAMILY[invariant as ExtendedCorpusInvariant] ?? CORPUS_INVARIANT_FAMILIES.core;

export const isGatingFailure = ({ invariant }: { invariant: CorpusInvariant }): boolean =>
  isGatingFamily(familyOf(invariant));

/** Milliseconds a single invariant may take on one file before the overrun is a finding. */
export const DEFAULT_INVARIANT_BUDGET_MS = 30_000;

export type CorpusInvariantInput = {
  /** The file as read, for the invariants that need the package rather than the model. */
  bytes: Uint8Array;
  /** The same bytes as a standalone buffer, which the save entry points require. */
  buffer: ArrayBuffer;
  parsed: Document;
  /** The main part as the package relationship names it. */
  documentPart: string;
  budgetMs: number;
};

/** Wall time per named stage, in milliseconds, for the performance census. */
export type StageTimings = Record<string, number>;

export type CorpusInvariantOutcome = {
  failures: CorpusFailure[];
  timings: StageTimings;
};

export type CorpusInvariantModule = {
  invariant: ExtendedCorpusInvariant;
  run: (input: CorpusInvariantInput) => Promise<CorpusInvariantOutcome>;
};

/**
 * Run one stage, recording its wall time whether it answers or throws.
 *
 * `run` may be synchronous: several stages are, and forcing them through an
 * `async` wrapper only to satisfy this signature would put a lint suppression
 * in every invariant.
 */
export const timeStage = async <T>(
  timings: StageTimings,
  stage: string,
  run: () => T | Promise<T>,
): Promise<T> => {
  const started = Bun.nanoseconds();
  try {
    return await run();
  } finally {
    timings[stage] = (Bun.nanoseconds() - started) / 1e6;
  }
};

export const totalStageMs = (timings: StageTimings): number =>
  Object.values(timings).reduce((total, ms) => total + ms, 0);
