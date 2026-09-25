import { panic, TaggedError } from "better-result";

/**
 * Process exit classes. The numbers match the stella command line's table so
 * a script driving both branches on one set of codes; folio uses the subset a
 * local file tool can produce.
 */
export const EXIT_CODES = {
  ok: 0,
  unexpected: 1,
  validation: 2,
  notFound: 6,
  permissionDenied: 8,
  conflict: 10,
} as const;

export type ExitCode = (typeof EXIT_CODES)[keyof typeof EXIT_CODES];

const EXIT_CODE_DESCRIPTIONS = {
  ok: "success",
  unexpected: "unexpected internal error",
  validation: "usage, input, or refused-operation error",
  notFound: "file, change, or comment not found",
  permissionDenied: "path outside the allowed roots, or unsafe to write through",
  conflict: "conflict with current state (stale version, lock held, destination exists)",
} as const satisfies Record<keyof typeof EXIT_CODES, string>;

/** Every exit code with its meaning, ordered numerically, for `--help`. */
export const exitCodeEntries = (): readonly { code: number; meaning: string }[] => {
  const descriptions: Record<string, string> = EXIT_CODE_DESCRIPTIONS;
  return Object.entries(EXIT_CODES)
    .toSorted(([, left], [, right]) => left - right)
    .map(([key, code]) => ({
      code,
      meaning: descriptions[key] ?? panic(`exit code ${key} has no description`),
    }));
};

/** Machine codes carried in `error.code` of a failure envelope. */
export const FOLIO_CLI_ERROR_CODES = {
  usage: "usage_error",
  invalidInput: "invalid_input",
  invalidDocument: "invalid_document",
  tooLarge: "too_large",
  notFound: "not_found",
  outsideRoot: "outside_root",
  unsafePath: "unsafe_path",
  invalidDestination: "invalid_destination",
  staleVersion: "stale_version",
  staleTarget: "stale_target",
  ambiguousTarget: "ambiguous_target",
  operationRejected: "operation_rejected",
  repackRequired: "repack_required",
  authorRequired: "author_required",
  locked: "locked",
  destinationExists: "destination_exists",
  transactionConflict: "transaction_conflict",
  integrityFailed: "integrity_failed",
  rendererUnavailable: "renderer_unavailable",
  internal: "internal_error",
} as const;

export type FolioCliErrorCode = (typeof FOLIO_CLI_ERROR_CODES)[keyof typeof FOLIO_CLI_ERROR_CODES];

const ERROR_CODE_EXIT = {
  usage_error: EXIT_CODES.validation,
  invalid_input: EXIT_CODES.validation,
  invalid_document: EXIT_CODES.validation,
  too_large: EXIT_CODES.validation,
  not_found: EXIT_CODES.notFound,
  outside_root: EXIT_CODES.permissionDenied,
  unsafe_path: EXIT_CODES.permissionDenied,
  invalid_destination: EXIT_CODES.validation,
  stale_version: EXIT_CODES.conflict,
  stale_target: EXIT_CODES.conflict,
  ambiguous_target: EXIT_CODES.validation,
  operation_rejected: EXIT_CODES.validation,
  repack_required: EXIT_CODES.validation,
  author_required: EXIT_CODES.validation,
  locked: EXIT_CODES.conflict,
  destination_exists: EXIT_CODES.conflict,
  transaction_conflict: EXIT_CODES.conflict,
  integrity_failed: EXIT_CODES.unexpected,
  renderer_unavailable: EXIT_CODES.validation,
  internal_error: EXIT_CODES.unexpected,
} as const satisfies Record<FolioCliErrorCode, ExitCode>;

export const exitCodeForError = (code: FolioCliErrorCode): ExitCode => ERROR_CODE_EXIT[code];

/**
 * Every expected failure of a command or tool call. `hint` names the next
 * step (re-read, pass a flag); `details` carries structured context such as
 * the skipped operations of a refused batch.
 */
export class FolioCliError extends TaggedError("FolioCliError")<{
  code: FolioCliErrorCode;
  message: string;
  hint?: string;
  details?: unknown;
}> {}

type CliErrorOptions = {
  code: FolioCliErrorCode;
  message: string;
  hint?: string | undefined;
  details?: unknown;
};

/** Build a {@link FolioCliError} without spelling absent optional fields. */
export const cliError = ({ code, message, hint, details }: CliErrorOptions): FolioCliError =>
  new FolioCliError({
    code,
    message,
    ...(hint !== undefined && { hint }),
    ...(details !== undefined && { details }),
  });
