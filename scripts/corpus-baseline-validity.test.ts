import { describe, expect, test } from "bun:test";

import {
  type CommittedCorpusFiles,
  corpusValidityIssues,
  loadCommittedCorpusFiles,
} from "./lib/corpus-baseline-validity";
import { EXTENDED_INVARIANT_FAMILY } from "./lib/corpus-invariants/contract";
import { CORPUS_INVARIANTS, failureSignature } from "./lib/corpus-signature";

const row = (invariant: string, message: string, files: number) => {
  const failure = { invariant, message, frame: "-" } as Parameters<typeof failureSignature>[0];
  return { ...failure, signature: failureSignature(failure), files };
};

const RESERIALIZE = "reserialize";

/**
 * The committed files as the checks see them, with one thing broken per case.
 * Nothing here reads a file: what is under test is which inconsistencies the
 * rules catch, not how the repository stores them.
 */
const committedWith = (override: Partial<CommittedCorpusFiles>): CommittedCorpusFiles =>
  ({
    baseline: {
      schemaVersion: 1,
      lockDigest: "t1-deadbeef",
      reportOnlyDigest: "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
      failedFiles: 1,
      entries: [],
    },
    families: new Map(),
    refusals: { schemaVersion: 1, entries: [] },
    dispositions: { schemaVersion: 1, entries: [] },
    reportOnly: { schemaVersion: 1, files: [] },
    lock: { schemaVersion: 1, manifestDigest: "-", fileCount: 0, totalBytes: 0, sources: [] },
    ...override,
  }) as CommittedCorpusFiles;

const detailsMatching = (files: CommittedCorpusFiles, pattern: RegExp): string[] =>
  corpusValidityIssues(files)
    .map(({ detail }) => detail)
    .filter((detail) => pattern.test(detail));

describe("the committed corpus files check out against themselves", () => {
  test("the repository's own files are consistent", async () => {
    expect(corpusValidityIssues(await loadCommittedCorpusFiles())).toEqual([]);
  });
});

/**
 * Each case breaks one thing that pull-request CI could not otherwise see: the
 * ratchet needs an hour of corpus to notice any of them, and a reviewer reading
 * a diff of a thousand rows will notice none.
 */
describe("what the check catches", () => {
  test("a row whose signature disagrees with its own fields", () => {
    const broken = { ...row(RESERIALIZE, "a difference", 3), signature: "hand | edited | -" };
    expect(
      detailsMatching(
        committedWith({ baseline: { ...committedWith({}).baseline, entries: [broken] } }),
        /disagrees with its own fields/u,
      ),
    ).toHaveLength(1);
  });

  test("duplicate and unsorted rows", () => {
    const first = row(RESERIALIZE, "aaa", 1);
    const second = row(RESERIALIZE, "bbb", 1);
    expect(
      detailsMatching(
        committedWith({
          baseline: { ...committedWith({}).baseline, entries: [second, first, first] },
        }),
        /not sorted|duplicate signature/u,
      ).length,
    ).toBeGreaterThan(1);
  });

  test("a row recording no files at all", () => {
    expect(
      detailsMatching(
        committedWith({
          baseline: { ...committedWith({}).baseline, entries: [row(RESERIALIZE, "a", 0)] },
        }),
        /records 0 files/u,
      ),
    ).toHaveLength(1);
  });

  test("a report-only family in a file that gates", () => {
    expect(
      detailsMatching(
        committedWith({
          baseline: {
            ...committedWith({}).baseline,
            entries: [row(EXTENDED_INVARIANT_FAMILY.performance, "slow", 1)],
          },
        }),
        /report-only family/u,
      ).length,
    ).toBeGreaterThan(0);
  });

  test("a signature that is both an expected refusal and a baseline defect", () => {
    const entry = row(CORPUS_INVARIANTS.parse, "refused", 1);
    expect(
      detailsMatching(
        committedWith({
          baseline: { ...committedWith({}).baseline, entries: [entry] },
          refusals: {
            schemaVersion: 1,
            entries: [{ signature: entry.signature, reason: "by design", files: 1 }],
          },
        }),
        /is also a baseline defect/u,
      ),
    ).toHaveLength(1);
  });

  test("a baseline measured over a lock the repository no longer carries", () => {
    expect(detailsMatching(committedWith({}), /the committed lock is/u).length).toBeGreaterThan(0);
  });

  test("a report-only exemption the lock cannot honour", () => {
    expect(
      detailsMatching(
        committedWith({
          reportOnly: {
            schemaVersion: 1,
            files: [
              { sourceId: "apache-poi", path: "gone.docx", sha256: "0".repeat(64), reason: "slow" },
            ],
          },
        }),
        /gone\.docx/u,
      ).length,
    ).toBeGreaterThan(0);
  });
});
