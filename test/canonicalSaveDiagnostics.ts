import type { SaveDiagnostic } from "../packages/core/src/docx/saveDiagnostics";

/** A successful full save reports that the selective writer declined the body. */
export const CANONICAL_SAVE_FALLBACK_DIAGNOSTIC = {
  type: "selectiveSaveRefused",
  part: "word/document.xml",
} as const satisfies SaveDiagnostic;

type SaveDiagnosticDisposition =
  | { status: "failure" }
  | { status: "permitted"; diagnostic: typeof CANONICAL_SAVE_FALLBACK_DIAGNOSTIC };

/** Every diagnostic kind requires a decision shared with the adapter contract. */
export const CANONICAL_SAVE_DIAGNOSTIC_DISPOSITIONS = {
  sourceReplayMismatch: { status: "failure" },
  selectiveSaveRefused: {
    status: "permitted",
    diagnostic: CANONICAL_SAVE_FALLBACK_DIAGNOSTIC,
  },
  sourceReplayUnavailable: { status: "failure" },
  canonicalResourceReplacement: { status: "failure" },
} as const satisfies Record<SaveDiagnostic["type"], SaveDiagnosticDisposition>;
