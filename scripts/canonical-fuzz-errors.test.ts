import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { failureMarker } from "../test/consumer-scenarios/support/failure-fingerprints";
import { CanonicalSessionRefusalError } from "../packages/core/src/controller/hiddenEditorManager";
import { CanonicalSaveDiagnosticError } from "../packages/core/src/docx/canonicalSave";
import type { SaveDiagnostic } from "../packages/core/src/docx/saveDiagnostics";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import {
  recordCanonicalFuzzError,
  isCanonicalSaveFallback,
  type CanonicalFuzzError,
  type CanonicalFuzzPhase,
} from "../tests/parity/canonicalFuzzErrors";
import {
  CanonicalBrowserOracleError,
  canonicalOracleFailureRecord,
  captureCanonicalOracleFailure,
} from "../tests/parity/canonicalOracleFailure";

const DIAGNOSTICS = {
  sourceReplayMismatch: { type: "sourceReplayMismatch", part: "word/document.xml" },
  selectiveSaveRefused: { type: "selectiveSaveRefused", part: "word/document.xml" },
  sourceReplayUnavailable: { type: "sourceReplayUnavailable", part: "word/document.xml" },
  canonicalResourceReplacement: {
    type: "canonicalResourceReplacement",
    gap: CANONICAL_GAP.resourceReplacement,
    part: "word/numbering.xml",
  },
} as const satisfies Record<SaveDiagnostic["type"], SaveDiagnostic>;

const PHASES = {
  load: { type: "load" },
  input: { type: "input", index: 0, action: "pasteListHtml" },
  undo: { type: "undo", index: 0, action: "pasteListHtml" },
  redo: { type: "redo", index: 0, action: "pasteListHtml" },
  save: { type: "save", index: 0, action: "pasteListHtml" },
  finalSave: { type: "finalSave" },
  reload: { type: "reload" },
} as const satisfies Record<CanonicalFuzzPhase["type"], CanonicalFuzzPhase>;

test(
  "save-error records retain every diagnostic and phase beyond assertion truncation",
  () => {
    const previousErrors = globalThis.__folioCanonicalFuzzErrors;
    const previousPhase = globalThis.__folioCanonicalFuzzPhase;
    try {
      assertProperty(
        fc.property(fc.string({ maxLength: 80 }), (message) => {
          for (const diagnostic of Object.values(DIAGNOSTICS)) {
            for (const phase of Object.values(PHASES)) {
              const errors: CanonicalFuzzError[] = [];
              globalThis.__folioCanonicalFuzzErrors = errors;
              globalThis.__folioCanonicalFuzzPhase = phase;
              const error = new CanonicalSaveDiagnosticError({
                message,
                gap: CANONICAL_GAP.save,
                diagnostic,
              });
              recordCanonicalFuzzError(error);
              recordCanonicalFuzzError(error);
              expect(errors).toHaveLength(2);
              expect(errors.at(0)).toEqual({
                status: "saveDiagnostic",
                type: "CanonicalSaveDiagnosticError",
                gap: CANONICAL_GAP.save,
                diagnostic,
                message,
                phase,
              });
              for (const recorded of errors) {
                expect(isCanonicalSaveFallback(recorded)).toBe(
                  diagnostic.type === "selectiveSaveRefused" &&
                    (phase.type === "save" || phase.type === "finalSave"),
                );
                expect(
                  isCanonicalSaveFallback({
                    message,
                    phase,
                    diagnostic: { ...diagnostic, part: "word/styles.xml" },
                    status: "saveDiagnostic",
                    type: "CanonicalSaveDiagnosticError",
                    gap: CANONICAL_GAP.save,
                  }),
                ).toBe(false);
              }
              const observations = [{ phase, errors }];
              const failure = new CanonicalBrowserOracleError({
                message: "Save phase failed",
                cause: new TypeError("x".repeat(5_000)),
                observations,
                errorCapture: { status: "complete" },
              });
              const marker = failureMarker({
                test: "canonical browser save reporting",
                seed: 11,
                path: "2:1:1:0:1:1",
                repro: "canonical seed 11",
                failure,
              });
              const record = canonicalOracleFailureRecord({ marker, failure, flow: [] });
              expect(record.error.length).toBeLessThan(5_000);
              expect(JSON.parse(JSON.stringify(record)).observations).toEqual(observations);
            }
          }
        }),
        { numRuns: 20 },
      );
    } finally {
      globalThis.__folioCanonicalFuzzErrors = previousErrors;
      globalThis.__folioCanonicalFuzzPhase = previousPhase;
    }
  },
  propertyTestTimeout(5_000),
);

test(
  "refusal records preserve every gap and phase without becoming save fallbacks",
  () => {
    const previousErrors = globalThis.__folioCanonicalFuzzErrors;
    const previousPhase = globalThis.__folioCanonicalFuzzPhase;
    try {
      assertProperty(
        fc.property(fc.string({ maxLength: 80 }), (message) => {
          for (const gap of Object.values(CANONICAL_GAP)) {
            for (const phase of [...Object.values(PHASES), undefined]) {
              const errors: CanonicalFuzzError[] = [];
              globalThis.__folioCanonicalFuzzErrors = errors;
              globalThis.__folioCanonicalFuzzPhase = phase;
              recordCanonicalFuzzError(new CanonicalSessionRefusalError({ message, gap }));
              expect(errors).toEqual([
                {
                  status: "refusal",
                  type: "CanonicalSessionRefusalError",
                  message,
                  gap,
                  phase: phase ?? null,
                },
              ]);
              expect(errors.some(isCanonicalSaveFallback)).toBe(false);
            }
          }
        }),
        { numRuns: 20 },
      );
    } finally {
      globalThis.__folioCanonicalFuzzErrors = previousErrors;
      globalThis.__folioCanonicalFuzzPhase = previousPhase;
    }
  },
  propertyTestTimeout(5_000),
);

test("unscoped ordinary errors retain their identity without becoming save diagnostics", () => {
  const previousErrors = globalThis.__folioCanonicalFuzzErrors;
  const previousPhase = globalThis.__folioCanonicalFuzzPhase;
  try {
    const errors: CanonicalFuzzError[] = [];
    globalThis.__folioCanonicalFuzzErrors = errors;
    globalThis.__folioCanonicalFuzzPhase = undefined;
    recordCanonicalFuzzError(new TypeError("unavailable"));
    expect(errors.some(isCanonicalSaveFallback)).toBe(false);
    expect(errors).toEqual([
      { status: "error", type: "TypeError", message: "unavailable", phase: null },
    ]);
  } finally {
    globalThis.__folioCanonicalFuzzErrors = previousErrors;
    globalThis.__folioCanonicalFuzzPhase = previousPhase;
  }
});

test("failure capture retains pending diagnostics and records an unavailable browser sink", async () => {
  const phase = PHASES.redo;
  const pending = {
    status: "saveDiagnostic",
    type: "CanonicalSaveDiagnosticError",
    gap: CANONICAL_GAP.save,
    diagnostic: DIAGNOSTICS.sourceReplayMismatch,
    message: "pending diagnostic",
    phase,
  } as const satisfies CanonicalFuzzError;
  const cause = new TypeError("Snapshot failed before draining errors");
  const observations = [{ phase, errors: [pending] }];
  const captured = await captureCanonicalOracleFailure({
    cause,
    observations,
    drainErrors: async () => [pending],
  });
  expect(captured.cause).toBe(cause);
  expect(captured.errorCapture).toEqual({ status: "complete" });
  expect(captured.observations).toEqual([{ phase, errors: [pending, pending] }]);
  const unavailable = await captureCanonicalOracleFailure({
    cause,
    observations,
    drainErrors: async () => {
      throw new TypeError("Page closed");
    },
  });
  expect(unavailable.cause).toBe(cause);
  expect(unavailable.errorCapture).toEqual({ status: "unavailable", message: "Page closed" });
  expect(unavailable.observations).toEqual(observations);
});
