import { expect, test } from "bun:test";

import {
  caseIdentity,
  compareBaseline,
  type EditBaseline,
  type EditCase,
  type EditReport,
} from "./corpus-edit-fuzz-ratchet.ts";

const original: EditCase = {
  document: "public-corpus/test-data/document/example.docx",
  sha256: "4409a5e95d229c5ab94fb24aac1a53ee47d0d66c9e75b7f39d80f10c218abedb",
  seed: 1141482985,
  class: "outcome",
  signature: "outcome:commentOnRange:comment",
  expected: "requested outcome after save and reopen",
  observed:
    "comment mismatch (diagnostic SHA-256 47295dfba12e0c86429afd2b446ba8c087d3d44e3a3bb37eb03d6400bcec7310)",
  operations: [
    {
      mode: "tracked-changes",
      operation: {
        type: "commentOnRange",
        range: {
          type: "textRange",
          story: "main",
          blockId: "42503136",
          startOffset: 43,
          endOffset: 47,
          selectedTextHash: "hyks5xo",
        },
        comment: { text: "Payment delivery delivery." },
      },
    },
  ],
};

const report = (cases: EditCase[], incomplete: EditReport["incomplete"] = []): EditReport => ({
  schemaVersion: 1,
  lockDigest: "pinned-lock",
  shard: null,
  sample: null,
  documents: 1,
  parsed: 1,
  attempts: 1,
  counts: {},
  cases,
  incomplete,
});

const baseline = (cases: EditCase[]): EditBaseline => ({
  schemaVersion: 2,
  lockDigest: "pinned-lock",
  identities: cases.map(caseIdentity),
  knownIncomplete: [],
});

test("a new failure cannot replace an old one under the same class and count", () => {
  const changed = {
    ...original,
    sha256: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
  };
  const result = compareBaseline([report([changed])], baseline([original]));
  expect(result.introduced).toEqual([caseIdentity(changed)]);
  expect(result.missing).toEqual([caseIdentity(original)]);
});

test("volatile SDK diagnostic hashes do not change a failure identity", () => {
  expect(caseIdentity({ ...original, observed: "another diagnostic SHA-256" })).toBe(
    caseIdentity(original),
  );
});

test("a duplicate case is not hidden by the identity set", () => {
  const result = compareBaseline([report([original, original])], baseline([original]));
  expect(result.duplicates).toEqual([caseIdentity(original)]);
});

test("unknown timeouts fail and exact known timeouts require later removal", () => {
  const timeout = {
    document: original.document,
    sha256: original.sha256,
    seed: original.seed,
    reason: "timeout" as const,
    detail: "120000ms deadline",
  };
  expect(
    compareBaseline([report([original], [timeout])], baseline([original])).unexpectedIncomplete,
  ).toEqual([timeout]);
  const known: EditBaseline = {
    ...baseline([original]),
    knownIncomplete: [{ ...timeout, finding: "CORPUS_EDIT_FUZZ_TIMEOUT" }],
  };
  expect(compareBaseline([report([original], [timeout])], known).unexpectedIncomplete).toEqual([]);
  expect(compareBaseline([report([original])], known).recoveredIncomplete).toEqual(
    known.knownIncomplete,
  );
});

test("a partial shard set cannot claim a full corpus result", () => {
  const first = { ...report([original]), shard: "1/4" };
  const second = { ...report([]), shard: "2/4" };
  expect(() => compareBaseline([first, second], baseline([original]))).toThrow(
    "every shard exactly once",
  );
});
