/** Stable identities for the failures in a pinned corpus edit-fuzz run. */
import { createHash } from "node:crypto";

import type { EditFailure } from "../../test/corpus-edit-fuzz-contract.ts";

export type EditCase = EditFailure & { document: string; sha256: string; seed: number };
export type Incomplete = {
  document: string;
  sha256: string;
  seed: number;
  reason: "timeout" | "worker";
  detail: string;
};
export type EditReport = {
  schemaVersion: 1;
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

export const FINDINGS = {
  CORPUS_EDIT_FUZZ_TIMEOUT:
    "The pinned document's edit worker exceeds its 120-second deadline; remove this disposition when it completes.",
} as const;

type KnownIncomplete = Omit<Incomplete, "reason"> & {
  reason: "timeout";
  finding: keyof typeof FINDINGS;
};
export type EditBaseline = {
  schemaVersion: 2;
  lockDigest: string;
  identities: string[];
  knownIncomplete: KnownIncomplete[];
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

const incompleteIdentity = ({ document, sha256, seed, reason, detail }: Incomplete): string =>
  JSON.stringify({ document, sha256, seed, reason, detail });

/** A sampled run may find no new case; a full run must match in both directions. */
export const compareBaseline = (
  reports: readonly EditReport[],
  baseline: EditBaseline,
): {
  introduced: string[];
  missing: string[];
  duplicates: string[];
  unexpectedIncomplete: Incomplete[];
  recoveredIncomplete: KnownIncomplete[];
} => {
  if (
    baseline.schemaVersion !== 2 ||
    reports.some(({ lockDigest }) => lockDigest !== baseline.lockDigest)
  ) {
    throw new Error("edit fuzz report and baseline versions or lock digests differ");
  }
  const fullRun = reports.length === 1 && reports[0]?.shard === null && reports[0]?.sample === null;
  const allShards =
    reports.length > 1 &&
    reports.every(({ sample }) => sample === null) &&
    Array.from({ length: reports.length }, (_, index) => `${index + 1}/${reports.length}`).every(
      (shard) => reports.some((report) => report.shard === shard),
    );
  if (!fullRun && !allShards && reports.length > 1) {
    throw new Error("edit fuzz check needs every shard exactly once");
  }
  const observed = reports.flatMap(({ cases }) => cases.map(caseIdentity));
  const incomplete = reports.flatMap(({ incomplete: cases }) => cases);
  const baselineIds = new Set(baseline.identities);
  const observedIds = new Set(observed);
  const knownIncomplete = new Set(baseline.knownIncomplete.map(incompleteIdentity));
  const observedIncomplete = new Set(incomplete.map(incompleteIdentity));
  const baselineDuplicates = duplicateIds(baseline.identities);
  if (baselineDuplicates.length > 0) {
    throw new Error(
      `edit fuzz baseline contains duplicate identities: ${baselineDuplicates.length}`,
    );
  }
  if (
    baseline.knownIncomplete.some(
      ({ reason, finding }) => reason !== "timeout" || !Object.hasOwn(FINDINGS, finding),
    ) ||
    duplicateIds(baseline.knownIncomplete.map(incompleteIdentity)).length > 0
  ) {
    throw new Error("edit fuzz baseline has an invalid or duplicate incomplete disposition");
  }
  return {
    introduced: [...observedIds].filter((id) => !baselineIds.has(id)).toSorted(),
    missing:
      fullRun || allShards ? [...baselineIds].filter((id) => !observedIds.has(id)).toSorted() : [],
    duplicates: duplicateIds(observed),
    unexpectedIncomplete: incomplete.filter(
      (item) => !knownIncomplete.has(incompleteIdentity(item)),
    ),
    recoveredIncomplete:
      fullRun || allShards
        ? baseline.knownIncomplete.filter(
            (item) => !observedIncomplete.has(incompleteIdentity(item)),
          )
        : [],
  };
};
