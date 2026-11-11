/** Stable identities for the failures in a pinned corpus edit-fuzz run. */
import { createHash } from "node:crypto";

import type { EditFailure } from "../../test/corpus-edit-fuzz-contract.ts";

export type EditCase = EditFailure & {
  document: string;
  sha256: string;
  seed: number;
};
export type Incomplete = {
  document: string;
  sha256: string;
  seed: number;
  reason: "timeout" | "worker";
  detail: string;
  /** Wall-clock time the worker ran before it ended or was stopped. */
  durationMs: number;
};
export type EditReport = {
  schemaVersion: 2;
  lockDigest: string;
  shard: string | null;
  sample: number | null;
  documents: number;
  parsed: number;
  attempts: number;
  counts: Record<string, number>;
  cases: EditCase[];
  incomplete: Incomplete[];
};

/**
 * The failure identities a pinned corpus run is expected to find.
 *
 * A timeout is no identity: a wall-clock limit drifts with machine load, so a
 * document near the deadline would pass one run and fail the next for nothing
 * the code did. Timeouts are listed in the run summary instead, and a timed
 * out document's own identities are not counted missing. A worker that
 * crashes always fails the ratchet.
 */
export type EditBaseline = {
  schemaVersion: 3;
  lockDigest: string;
  identities: string[];
};

/** Pinned document hash, failure category and reduced request; diagnostics may contain volatile SDK ids. */
export const caseIdentity = ({
  sha256,
  seed,
  signature,
  expected,
  operations,
}: EditCase): string => {
  const detail = createHash("sha256")
    .update(JSON.stringify({ seed, expected, operations }))
    .digest("hex");
  return `${sha256}::${signature}::${detail}`;
};

const duplicateIds = (ids: readonly string[]): string[] => {
  const seen = new Set<string>();
  const duplicates = new Set<string>();
  for (const id of ids) {
    if (seen.has(id)) duplicates.add(id);
    seen.add(id);
  }
  return [...duplicates].toSorted();
};

const documentOf = (identity: string): string =>
  identity.slice(0, identity.indexOf("::"));

/** A sampled run may find no new case; a full run must match in both directions. */
export const compareBaseline = (
  reports: readonly EditReport[],
  baseline: EditBaseline,
): {
  introduced: string[];
  missing: string[];
  duplicates: string[];
  crashed: Incomplete[];
} => {
  if (
    baseline.schemaVersion !== 3 ||
    reports.some(
      ({ schemaVersion, lockDigest }) =>
        schemaVersion !== 2 || lockDigest !== baseline.lockDigest,
    )
  ) {
    throw new Error(
      "edit fuzz report and baseline versions or lock digests differ",
    );
  }
  const fullRun =
    reports.length === 1 &&
    reports[0]?.shard === null &&
    reports[0]?.sample === null;
  const allShards =
    reports.length > 1 &&
    reports.every(({ sample }) => sample === null) &&
    Array.from(
      { length: reports.length },
      (_, index) => `${index + 1}/${reports.length}`,
    ).every((shard) => reports.some((report) => report.shard === shard));
  if (!fullRun && !allShards && reports.length > 1) {
    throw new Error("edit fuzz check needs every shard exactly once");
  }
  const observed = reports.flatMap(({ cases }) => cases.map(caseIdentity));
  const incomplete = reports.flatMap(({ incomplete: cases }) => cases);
  const timedOut = new Set(
    incomplete
      .filter(({ reason }) => reason === "timeout")
      .map(({ sha256 }) => sha256),
  );
  const baselineIds = new Set(baseline.identities);
  const observedIds = new Set(observed);
  const baselineDuplicates = duplicateIds(baseline.identities);
  if (baselineDuplicates.length > 0) {
    throw new Error(
      `edit fuzz baseline contains duplicate identities: ${baselineDuplicates.length}`,
    );
  }
  return {
    introduced: [...observedIds]
      .filter((id) => !baselineIds.has(id))
      .toSorted(),
    missing:
      fullRun || allShards
        ? [...baselineIds]
            .filter(
              (id) => !observedIds.has(id) && !timedOut.has(documentOf(id)),
            )
            .toSorted()
        : [],
    duplicates: duplicateIds(observed),
    crashed: incomplete.filter(({ reason }) => reason === "worker"),
  };
};

/** The baseline a complete run records: its identities, once each, sorted. */
export const baselineOf = (reports: readonly EditReport[]): EditBaseline => {
  const lockDigest = reports.at(0)?.lockDigest;
  if (
    lockDigest === undefined ||
    reports.some((report) => report.lockDigest !== lockDigest)
  )
    throw new Error("edit fuzz reports name different corpus locks");
  return {
    schemaVersion: 3,
    lockDigest,
    identities: [
      ...new Set(reports.flatMap(({ cases }) => cases.map(caseIdentity))),
    ].toSorted(),
  };
};
