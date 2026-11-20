/**
 * Tools that rewrite a package's XML parts directly rather than through the
 * document model: each runs against the package in memory and returns a
 * report. Without a destination the report is all that comes back (a dry
 * run of exactly what a write would do); with one, the touched parts are
 * written back and the package commits like any other change.
 */

import { panic, Result } from "better-result";

import { ensureParaIds, FolioDocxReviewer } from "@stll/folio-core/server";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import { fillTemplate } from "./fill-template";
import { rewriteNumbering } from "./numbering";
import { type OoxmlPackage, openOoxmlPackage } from "./ooxml-package";
import { setProofingLanguage } from "./proofing-language";
import { sanitizePackage } from "./sanitize";

type RewriteArgs = Readonly<Record<string, unknown>>;

/**
 * A rewrite of the package alone, or one that also reads the document
 * model: that one runs on the package with a `w14:paraId` on every
 * paragraph, so the block ids the model reports name paragraphs in the XML.
 */
type PackageRewrite =
  | {
      type: "package";
      run: (
        pkg: OoxmlPackage,
        args: RewriteArgs,
      ) => Result<unknown, FolioCliError>;
    }
  | {
      type: "withDocument";
      run: (
        pkg: OoxmlPackage,
        args: RewriteArgs,
        reviewer: FolioDocxReviewer,
      ) => Result<unknown, FolioCliError>;
    };

/** Every package rewrite, by the name a registry entry gives it. */
const PACKAGE_REWRITES = {
  sanitize: { type: "package", run: sanitizePackage },
  proofingLanguage: { type: "package", run: setProofingLanguage },
  numbering: { type: "withDocument", run: rewriteNumbering },
  fillTemplate: { type: "withDocument", run: fillTemplate },
} as const satisfies Record<string, PackageRewrite>;

export type PackageRewriteName = keyof typeof PACKAGE_REWRITES;

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

/** The bytes with an id on every paragraph, and the model read from them. */
const withDocument = async (
  bytes: Uint8Array,
): Promise<
  Result<{ bytes: Uint8Array; reviewer: FolioDocxReviewer }, FolioCliError>
> => {
  const normalized = await Result.tryPromise({
    try: () => ensureParaIds(bytes),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `This command needs a paragraph id on every paragraph, which the package does not allow: ${describe(error)}`,
      }),
  });
  if (normalized.isErr()) return Result.err(normalized.error);
  const { docx } = normalized.value;
  const reviewer = await Result.tryPromise({
    try: () =>
      FolioDocxReviewer.fromBuffer(docx.slice().buffer, { session: "legacy" }),
    catch: (error) =>
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: `The document could not be read: ${describe(error)}`,
      }),
  });
  if (reviewer.isErr()) return Result.err(reviewer.error);
  return Result.ok({ bytes: docx, reviewer: reviewer.value });
};

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
  const rewrite: PackageRewrite = PACKAGE_REWRITES[name];
  const document =
    rewrite.type === "withDocument" ? await withDocument(bytes) : null;
  if (document?.isErr() === true) return Result.err(document.error);
  const opened = await openOoxmlPackage(
    document?.isOk() === true ? document.value.bytes : bytes,
  );
  if (opened.isErr()) return Result.err(opened.error);
  const pkg = opened.value;
  let report: Result<unknown, FolioCliError>;
  switch (rewrite.type) {
    case "package":
      report = rewrite.run(pkg, args);
      break;
    case "withDocument":
      report =
        document?.isOk() === true
          ? rewrite.run(pkg, args, document.value.reviewer)
          : panic("A document rewrite ran without its document");
      break;
    default: {
      const unreachable: never = rewrite;
      return panic("Unhandled package rewrite", { unreachable });
    }
  }
  if (report.isErr()) return Result.err(report.error);
  const parts = pkg.changedParts();
  return Result.ok({
    report: report.value,
    parts,
    bytes: async () => (parts.length === 0 ? null : await pkg.toBytes()),
  });
};

/** The refusal a write gets when the rewrite or pass found nothing to change. */
export const nothingToRewrite = (tool: string): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.operationRejected,
    message: `${tool} found nothing to change; nothing was written.`,
    hint: "Run it without --in-place or -o to see the report.",
  });
