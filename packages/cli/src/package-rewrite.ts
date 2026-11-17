/**
 * Tools that rewrite a package's XML parts directly rather than through the
 * document model: each runs against the package in memory and returns a
 * report. Without a destination the report is all that comes back (a dry
 * run of exactly what a write would do); with one, the touched parts are
 * written back and the package commits like any other change.
 */

import { Result } from "better-result";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { type OoxmlPackage, openOoxmlPackage } from "./ooxml-package";
import { setProofingLanguage } from "./proofing-language";
import { sanitizePackage } from "./sanitize";

type PackageRewrite = (
  pkg: OoxmlPackage,
  args: Readonly<Record<string, unknown>>,
) => Result<unknown, FolioCliError>;

/** Every package rewrite, by the name a registry entry gives it. */
const PACKAGE_REWRITES = {
  sanitize: sanitizePackage,
  proofingLanguage: setProofingLanguage,
} as const satisfies Record<string, PackageRewrite>;

export type PackageRewriteName = keyof typeof PACKAGE_REWRITES;

export type PackageRewriteOutcome = {
  /** What the rewrite found and changed (or, as a dry run, would change). */
  report: unknown;
  /** Parts it changed or removed, sorted. */
  parts: readonly string[];
  /** The rewritten package; `null` when nothing changed. */
  bytes: () => Promise<Uint8Array<ArrayBuffer> | null>;
};

/** Run one rewrite against a package's bytes. */
export const runPackageRewrite = async (
  name: PackageRewriteName,
  bytes: Uint8Array,
  args: Readonly<Record<string, unknown>>,
): Promise<Result<PackageRewriteOutcome, FolioCliError>> => {
  const opened = await openOoxmlPackage(bytes);
  if (opened.isErr()) return Result.err(opened.error);
  const pkg = opened.value;
  const report = PACKAGE_REWRITES[name](pkg, args);
  if (report.isErr()) return Result.err(report.error);
  const parts = pkg.changedParts();
  return Result.ok({
    report: report.value,
    parts,
    bytes: async () => (parts.length === 0 ? null : await pkg.toBytes()),
  });
};

/** The refusal a write gets when the rewrite found nothing to change. */
export const nothingToRewrite = (tool: string): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.operationRejected,
    message: `${tool} found nothing to change; nothing was written.`,
    hint: "Run it without --in-place or -o to see the report.",
  });
