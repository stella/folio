import { describe, expect, test } from "bun:test";

import { editorRoundTripVerdict } from "./corpus-editor-round-trip";
import { CORPUS_EVIDENCE } from "./lib/corpus-census";
import type { CorpusCheckResult } from "./lib/corpus-check";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import type { CorpusFailure } from "./lib/corpus-signature";

type CheckedResult = Extract<CorpusCheckResult, { kind: "checked" }>;

const checkedResult = (
  overrides: Partial<Omit<CheckedResult, "kind" | "cost" | "producer">> = {},
) =>
  ({
    kind: "checked",
    failures: [],
    declaredRefusals: [],
    producer: "word/16",
    cost: {
      bytes: 128,
      parseMs: 1,
      peakRssBytes: 1024,
      referenceMs: 1,
    },
    timings: { "editor-round-trip.to-prose": 1 },
    ...overrides,
  }) satisfies CheckedResult;

const editorFailure = {
  invariant: EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
  message: "editor round trip changed a field",
  frame: "-",
} satisfies CorpusFailure;

describe("the scoped editor round-trip verdict", () => {
  test("a completed clean editor round trip is gating evidence", () => {
    expect(editorRoundTripVerdict(checkedResult(), false)).toEqual({
      kind: "checked",
      status: "complete",
      evidence: CORPUS_EVIDENCE.gating,
      editorFailures: [],
    });
  });

  test("a completed editor round-trip failure is preserved", () => {
    expect(
      editorRoundTripVerdict(
        checkedResult({ failures: [editorFailure] }),
        false,
      ),
    ).toEqual({
      kind: "checked",
      status: "complete",
      evidence: CORPUS_EVIDENCE.gating,
      editorFailures: [editorFailure],
    });
  });

  test("a listed report-only file stays report-only evidence", () => {
    expect(
      editorRoundTripVerdict(
        checkedResult({ failures: [editorFailure] }),
        true,
      ),
    ).toEqual({
      kind: "checked",
      status: "complete",
      evidence: CORPUS_EVIDENCE.reportOnly,
      editorFailures: [editorFailure],
    });
  });

  test("a truncated editor check is incomplete even with a result", () => {
    expect(
      editorRoundTripVerdict(
        checkedResult({ truncatedAt: "editor-round-trip.compare" }),
        false,
      ),
    ).toEqual({
      kind: "checked",
      status: "incomplete",
      evidence: CORPUS_EVIDENCE.reportOnly,
      editorFailures: [],
      reason: "Truncated at editor-round-trip.compare",
    });
  });

  test("a checked parse failure without editor timings is incomplete", () => {
    const parseFailure = {
      invariant: "parse",
      message: "document could not be parsed",
      frame: "-",
    } satisfies CorpusFailure;
    expect(
      editorRoundTripVerdict(
        checkedResult({ failures: [parseFailure], timings: {} }),
        false,
      ),
    ).toEqual({
      kind: "checked",
      status: "incomplete",
      evidence: CORPUS_EVIDENCE.reportOnly,
      editorFailures: [],
      reason: "Editor round-trip invariant did not run",
    });
  });

  test("a classified non-DOCX is a completed skip", () => {
    const notADocx = {
      kind: "not-a-docx",
      reason: "not-a-zip",
      detail: "not a ZIP archive",
    } satisfies CorpusCheckResult;
    expect(editorRoundTripVerdict(notADocx, false)).toEqual({
      kind: "not-a-docx",
      status: "complete",
      evidence: CORPUS_EVIDENCE.reportOnly,
      editorFailures: [],
      notADocx: "not-a-zip",
    });
  });
});
