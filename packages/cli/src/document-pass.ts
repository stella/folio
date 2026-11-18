/**
 * Tools that run over the whole opened document rather than a batch of
 * caller-addressed operations: each inspects the document and, given a
 * destination, changes it through the reviewer, as tracked changes unless
 * the caller asks for direct edits. Without a destination only the
 * inspection runs and nothing is written.
 */

import { Result } from "better-result";

import type {
  FolioAIEditApplyMode,
  FolioDocxReviewer,
} from "@stll/folio-core/server";

import type { FolioCliError } from "./errors";
import { fillPlaceholders, inspectPlaceholders } from "./placeholders";

/** What a pass applies with: the edit mode and the revision stamp. */
export type DocumentPassContext = {
  mode: FolioAIEditApplyMode;
  date: string;
  idSeed: number;
};

type DocumentPass = {
  inspect: (
    reviewer: FolioDocxReviewer,
    args: Readonly<Record<string, unknown>>,
  ) => Result<unknown, FolioCliError>;
  apply: (
    reviewer: FolioDocxReviewer,
    args: Readonly<Record<string, unknown>>,
    context: DocumentPassContext,
  ) => Result<DocumentPassOutcome, FolioCliError>;
};

/** What an applied pass reports, and whether it changed the document. */
export type DocumentPassOutcome = { result: unknown; changed: boolean };

const DOCUMENT_PASSES = {
  placeholders: {
    inspect: inspectPlaceholders,
    apply: (reviewer, args, context) =>
      fillPlaceholders(reviewer, args, context).map((result) => ({
        result,
        changed: result.filled > 0,
      })),
  },
} as const satisfies Record<string, DocumentPass>;

export type DocumentPassName = keyof typeof DOCUMENT_PASSES;

/** Inspect without changing anything. */
export const inspectDocument = (
  name: DocumentPassName,
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
): Result<unknown, FolioCliError> =>
  DOCUMENT_PASSES[name].inspect(reviewer, args);

/** Apply a pass to the reviewer; the caller saves it. */
export const applyDocumentPass = (
  name: DocumentPassName,
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
  context: DocumentPassContext,
): Result<DocumentPassOutcome, FolioCliError> =>
  DOCUMENT_PASSES[name].apply(reviewer, args, context);
