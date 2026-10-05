import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../test/property-testing";
import { failureMarker } from "../../test/consumer-scenarios/support/failure-fingerprints";
import { CanonicalSaveDiagnosticError } from "../../packages/core/src/docx/canonicalSave";
import type { SaveDiagnostic } from "../../packages/core/src/docx/saveDiagnostics";
import { CANONICAL_GAP } from "../../packages/core/src/types/canonicalCapabilities";
import {
  recordCanonicalFuzzError,
  type CanonicalFuzzError,
  type CanonicalFuzzPhase,
} from "./canonicalFuzzErrors";
import {
  CanonicalBrowserOracleError,
  canonicalOracleFailureRecord,
} from "./canonicalOracleFailure";

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
              const observations = [{ phase, errors }];
              const failure = new CanonicalBrowserOracleError({
                message: "Save phase failed",
                cause: new TypeError("x".repeat(5_000)),
                observations,
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

test("unscoped ordinary errors retain their identity without becoming save diagnostics", () => {
  const previousErrors = globalThis.__folioCanonicalFuzzErrors;
  const previousPhase = globalThis.__folioCanonicalFuzzPhase;
  try {
    const errors: CanonicalFuzzError[] = [];
    globalThis.__folioCanonicalFuzzErrors = errors;
    globalThis.__folioCanonicalFuzzPhase = undefined;
    recordCanonicalFuzzError(new TypeError("unavailable"));
    expect(errors).toEqual([
      { status: "error", type: "TypeError", message: "unavailable", phase: null },
    ]);
  } finally {
    globalThis.__folioCanonicalFuzzErrors = previousErrors;
    globalThis.__folioCanonicalFuzzPhase = previousPhase;
  }
});
