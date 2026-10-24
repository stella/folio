import { describe, expect, test } from "bun:test";

import { editorRoundTripVerdict } from "./corpus-editor-round-trip";
import { CORPUS_EVIDENCE } from "./lib/corpus-census";
import type { CorpusCheckResult } from "./lib/corpus-check";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import { CORPUS_INVARIANTS, type CorpusFailure } from "./lib/corpus-signature";

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
      editorRun: "evaluated",
      evidence: CORPUS_EVIDENCE.gating,
      parseFailures: [],
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
      editorRun: "evaluated",
      evidence: CORPUS_EVIDENCE.gating,
      parseFailures: [],
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
      editorRun: "evaluated",
      evidence: CORPUS_EVIDENCE.reportOnly,
      parseFailures: [],
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
      editorRun: "evaluated",
      evidence: CORPUS_EVIDENCE.reportOnly,
      parseFailures: [],
      editorFailures: [],
      reason: "Truncated at editor-round-trip.compare",
    });
  });

  test("truncation stays incomplete even when parsing also failed", () => {
    const parseFailure = {
      invariant: CORPUS_INVARIANTS.parse,
      message: "archive exceeded the parser size limit",
      frame: "-",
    } satisfies CorpusFailure;
    expect(
      editorRoundTripVerdict(
        checkedResult({
          failures: [parseFailure],
          timings: {},
          truncatedAt: "parse.preflight",
        }),
        false,
      ),
    ).toEqual({
      kind: "checked",
      status: "incomplete",
      editorRun: "not-run",
      evidence: CORPUS_EVIDENCE.reportOnly,
      parseFailures: [parseFailure],
      editorFailures: [],
      reason: "Truncated at parse.preflight",
    });
  });

  test("an actual parser refusal is a complete non-editor outcome", () => {
    const parseFailure = {
      invariant: CORPUS_INVARIANTS.parse,
      message: "parser refused a malformed document",
      frame: "-",
    } satisfies CorpusFailure;
    expect(
      editorRoundTripVerdict(
        checkedResult({ failures: [parseFailure], timings: {} }),
        false,
      ),
    ).toEqual({
      kind: "parse-failed",
      status: "complete",
      editorRun: "not-run",
      evidence: CORPUS_EVIDENCE.gating,
      parseFailures: [parseFailure],
      editorFailures: [],
    });
  });

  test("an unexplained missing editor run is incomplete", () => {
    expect(
      editorRoundTripVerdict(checkedResult({ timings: {} }), false),
    ).toEqual({
      kind: "incomplete",
      status: "incomplete",
      editorRun: "not-run",
      evidence: CORPUS_EVIDENCE.gating,
      parseFailures: [],
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
      editorRun: "not-run",
      evidence: CORPUS_EVIDENCE.reportOnly,
      parseFailures: [],
      editorFailures: [],
      notADocx: "not-a-zip",
    });
  });
});
