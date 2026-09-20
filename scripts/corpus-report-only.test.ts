/**
 * The committed list is the only way out of the gating set.
 *
 * Whether a file contributes gating evidence used to depend on whether it
 * finished inside a budget, which depends on how loaded the runner was. These
 * tests hold the replacement rule: a listed file never gates, finished or not;
 * an unlisted file that stops degrades the run; an entry the corpus no longer
 * carries fails the gate; and a baseline written before a listing is not
 * ratcheted down to it.
 */

import { describe, expect, test } from "bun:test";

import {
  baselineFromCensus,
  compareToBaseline,
  isFailingViolation,
  reportOnlyListChanged,
} from "./lib/corpus-baseline";
import { CensusBuilder, type CorpusCensus, type CorpusFileId } from "./lib/corpus-census";
import { compareFamilyToBaseline, familyBaselineFromCensus } from "./lib/corpus-family-baseline";
import { CORPUS_EVIDENCE, evidenceOf } from "./lib/corpus-census";
import { FamilyCensusBuilder } from "./lib/corpus-family-census";
import {
  CORPUS_INVARIANT_FAMILIES,
  EXTENDED_CORPUS_INVARIANTS,
} from "./lib/corpus-invariants/contract";
import type { CorpusLock } from "./lib/corpus-manifest";
import {
  type ReportOnlyFiles,
  deadReportOnlyEntries,
  loadReportOnlyFiles,
  reportOnlyFileIds,
  reportOnlyFilesDigest,
  validateReportOnlyFiles,
} from "./lib/corpus-report-only";
import { CORPUS_INVARIANTS, failureFromAssertion } from "./lib/corpus-signature";

const LOCK_DIGEST = "c".repeat(64);
const SHA = (seed: string): string => seed.repeat(64).slice(0, 64);

const file = (name: string): CorpusFileId => ({
  sourceId: "source",
  path: `${name}.docx`,
  sha256: SHA(name),
});

const GATING = failureFromAssertion(CORPUS_INVARIANTS.fixedPoint, "visible text changed");
const TIMING = failureFromAssertion(
  EXTENDED_CORPUS_INVARIANTS.performance,
  "reserialize exceeded its per-file time budget",
);

const listOf = (...names: readonly string[]): ReportOnlyFiles => ({
  schemaVersion: 1,
  files: names.map((name) => ({
    sourceId: "source",
    path: `${name}.docx`,
    sha256: SHA(name),
    reason: `${name} is too slow to gate on`,
  })),
});

const lockOf = (...names: readonly string[]): CorpusLock => ({
  schemaVersion: 1,
  manifestDigest: LOCK_DIGEST,
  fileCount: names.length,
  totalBytes: names.length,
  sources: [
    {
      id: "source",
      commit: "0".repeat(40),
      tier: 1,
      files: names.map((name) => ({ path: `${name}.docx`, sha256: SHA(name), bytes: 1 })),
    },
  ],
});

describe("a listed file is measured and never gated", () => {
  const censusWith = (kind: "complete" | "report-only"): CorpusCensus => {
    const builder = new CensusBuilder(LOCK_DIGEST, reportOnlyFilesDigest(listOf("b")));
    builder.add(file("a"), { kind: "complete", failures: [GATING] });
    builder.add(
      file("b"),
      kind === "complete"
        ? { kind: "complete", failures: [GATING, TIMING] }
        : { kind: "report-only", failures: [GATING, TIMING] },
    );
    return builder.build();
  };

  test("its gating findings do not reach the census, even when it finished", () => {
    const signature = censusWith("report-only").signatures.find(
      (entry) => entry.invariant === CORPUS_INVARIANTS.fixedPoint,
    );
    expect(signature?.files).toBe(1);
    expect(signature?.examples.map((example) => example.path)).toEqual(["a.docx"]);
  });

  test("the same file would have gated had it not been listed", () => {
    const signature = censusWith("complete").signatures.find(
      (entry) => entry.invariant === CORPUS_INVARIANTS.fixedPoint,
    );
    expect(signature?.files).toBe(2);
  });

  test("its timing findings do reach the census", () => {
    const signature = censusWith("report-only").signatures.find(
      (entry) => entry.invariant === EXTENDED_CORPUS_INVARIANTS.performance,
    );
    expect(signature?.files).toBe(1);
  });

  test("it is counted, and neither passes nor fails", () => {
    const census = censusWith("report-only");
    expect(census.files).toBe(2);
    expect(census.reportOnly).toBe(1);
    expect(census.truncated).toBe(0);
    expect(census.passed).toBe(0);
    expect(census.failedFiles).toBe(1);
  });

  test("the family census drops the same findings, from the same decision", () => {
    const listed = { kind: "report-only" as const, failures: [GATING, TIMING] };
    const builder = new FamilyCensusBuilder(LOCK_DIGEST, reportOnlyFilesDigest(listOf("b")));
    builder.add({
      file: file("b"),
      bytes: 1024,
      parseMs: 1,
      peakRssBytes: 1,
      referenceMs: 1,
      producer: "word/16",
      failures: [
        failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.reserialize, "a serializer difference"),
        TIMING,
      ],
      timings: {},
      evidence: evidenceOf(listed),
    });
    const census = builder.build();
    expect(evidenceOf(listed)).toBe(CORPUS_EVIDENCE.reportOnly);
    expect(census.signatures.map((entry) => entry.family)).toEqual([
      CORPUS_INVARIANT_FAMILIES.performance,
    ]);
  });
});

describe("a baseline written before a listing is not ratcheted down to it", () => {
  const before = (): CorpusCensus => {
    const builder = new CensusBuilder(LOCK_DIGEST, reportOnlyFilesDigest(listOf()));
    builder.add(file("a"), { kind: "complete", failures: [GATING] });
    builder.add(file("b"), { kind: "complete", failures: [GATING] });
    return builder.build();
  };

  const after = (extra: readonly CorpusFileId[] = []): CorpusCensus => {
    const builder = new CensusBuilder(LOCK_DIGEST, reportOnlyFilesDigest(listOf("b")));
    builder.add(file("a"), { kind: "complete", failures: [GATING] });
    builder.add(file("b"), { kind: "report-only", failures: [GATING] });
    for (const id of extra) {
      builder.add(id, { kind: "complete", failures: [GATING] });
    }
    return builder.build();
  };

  test("the two runs disagree about the list", () => {
    expect(reportOnlyListChanged(baselineFromCensus(before()), after())).toBe(true);
    expect(reportOnlyListChanged(baselineFromCensus(after()), after())).toBe(false);
  });

  test("a signature that lost the listed file is kept, not resolved", () => {
    const violations = compareToBaseline(baselineFromCensus(before()), after());
    expect(violations.map((violation) => violation.kind)).toEqual(["report-only-list-changed"]);
    expect(violations.every((violation) => !isFailingViolation(violation))).toBe(true);
  });

  test("a listing cannot hide a regression: more files still fails", () => {
    const violations = compareToBaseline(
      baselineFromCensus(before()),
      after([file("c"), file("d")]),
    );
    expect(violations.find(isFailingViolation)?.kind).toBe("more-files");
  });

  test("once re-measured, a real shrink ratchets again", () => {
    const builder = new CensusBuilder(LOCK_DIGEST, reportOnlyFilesDigest(listOf("b")));
    builder.add(file("b"), { kind: "report-only", failures: [GATING] });
    const violations = compareToBaseline(baselineFromCensus(after()), builder.build());
    expect(violations.map((violation) => violation.kind)).toEqual(["resolved-signature"]);
    expect(violations.every(isFailingViolation)).toBe(true);
  });

  test("a family baseline follows the same rule", () => {
    const familyCensus = (
      digest: string,
      evidence: typeof CORPUS_EVIDENCE.gating | typeof CORPUS_EVIDENCE.reportOnly,
    ) => {
      const builder = new FamilyCensusBuilder(LOCK_DIGEST, digest);
      builder.add({
        file: file("b"),
        bytes: 1024,
        parseMs: 1,
        peakRssBytes: 1,
        referenceMs: 1,
        producer: "word/16",
        failures: [
          failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.reserialize, "a serializer difference"),
        ],
        timings: {},
        evidence,
      });
      return builder.build();
    };
    const baseline = familyBaselineFromCensus(
      familyCensus(reportOnlyFilesDigest(listOf()), CORPUS_EVIDENCE.gating),
      CORPUS_INVARIANT_FAMILIES.reserialize,
    );
    const violations = compareFamilyToBaseline(
      baseline,
      familyCensus(reportOnlyFilesDigest(listOf("b")), CORPUS_EVIDENCE.reportOnly),
    );
    expect(violations.map((violation) => violation.kind)).toEqual(["report-only-list-changed"]);
    expect(violations.every((violation) => !isFailingViolation(violation))).toBe(true);
  });
});

describe("the list may not rot", () => {
  test("an entry the corpus no longer carries is reported", () => {
    expect(deadReportOnlyEntries(listOf("a"), lockOf("b"))).toEqual([
      "source/a.docx: no such file in corpus/sources.lock.json",
    ]);
  });

  test("an entry whose bytes changed is reported", () => {
    const repinned: ReportOnlyFiles = {
      schemaVersion: 1,
      files: [{ sourceId: "source", path: "a.docx", sha256: SHA("z"), reason: "still slow" }],
    };
    expect(deadReportOnlyEntries(repinned, lockOf("a"))).toHaveLength(1);
  });

  test("a live entry is not", () => {
    expect(deadReportOnlyEntries(listOf("a"), lockOf("a", "b"))).toEqual([]);
  });
});

describe("the list's shape", () => {
  const entry = {
    sourceId: "source",
    path: "a.docx",
    sha256: SHA("a"),
    reason: "too slow to gate on",
  };

  test("a valid list has no issues", () => {
    expect(validateReportOnlyFiles({ schemaVersion: 1, files: [entry] })).toEqual([]);
  });

  test("an entry without a reason is rejected", () => {
    const { reason: _reason, ...withoutReason } = entry;
    expect(validateReportOnlyFiles({ schemaVersion: 1, files: [withoutReason] })).toContain(
      "files[0].reason: expected a non-empty string",
    );
    expect(
      validateReportOnlyFiles({ schemaVersion: 1, files: [{ ...entry, reason: "" }] }),
    ).toHaveLength(1);
  });

  test("an entry without content identity is rejected", () => {
    expect(
      validateReportOnlyFiles({ schemaVersion: 1, files: [{ ...entry, sha256: "short" }] }),
    ).toContain("files[0].sha256: expected a lowercase SHA-256");
  });

  test("an unsorted or duplicated list is rejected", () => {
    const second = { ...entry, path: "b.docx", sha256: SHA("b") };
    expect(validateReportOnlyFiles({ schemaVersion: 1, files: [second, entry] })).toContain(
      "files: entries must be sorted by source id, then path",
    );
    expect(validateReportOnlyFiles({ schemaVersion: 1, files: [entry, entry] })).toContain(
      "files: the same file is listed twice",
    );
  });

  test("the digest follows the set, not the prose", () => {
    const reworded: ReportOnlyFiles = {
      schemaVersion: 1,
      files: [{ sourceId: "source", path: "a.docx", sha256: SHA("a"), reason: "rewritten" }],
    };
    expect(reportOnlyFilesDigest(reworded)).toBe(reportOnlyFilesDigest(listOf("a")));
    expect(reportOnlyFilesDigest(listOf("a", "b"))).not.toBe(reportOnlyFilesDigest(listOf("a")));
  });
});

describe("the committed list", () => {
  test("is valid, and every entry names a file the lock carries", async () => {
    const list = await loadReportOnlyFiles();
    const lock = (await Bun.file("corpus/sources.lock.json").json()) as CorpusLock;
    expect(validateReportOnlyFiles(list)).toEqual([]);
    expect(deadReportOnlyEntries(list, lock)).toEqual([]);
    expect(list.files.length).toBeGreaterThan(0);
  });

  test("every entry states a reason", async () => {
    const list = await loadReportOnlyFiles();
    for (const entry of list.files) {
      expect(entry.reason.length).toBeGreaterThan(20);
    }
  });

  test("its ids are the ids a run matches files by", async () => {
    const ids = reportOnlyFileIds(await loadReportOnlyFiles());
    expect(ids.has("officeparser/test/files/oversized_archive.docx")).toBe(true);
  });
});
