/**
 * Allowed roots for the MCP server: every path a tool call reads or writes
 * must resolve, through symlinks, inside one of them. A destination that does
 * not exist yet is checked through its parent directory.
 */

import { Result } from "better-result";
import { realpath, stat } from "node:fs/promises";
import path from "node:path";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

/** Real paths of the allowed root directories. */
export type AllowedRoots = readonly string[];

/** Resolve each root to its real path; every root must be an existing directory. */
export const resolveRoots = async (
  roots: readonly string[],
): Promise<Result<AllowedRoots, FolioCliError>> => {
  const resolved: string[] = [];
  for (const root of roots) {
    const real = await Result.tryPromise(() => realpath(path.resolve(root)));
    const info = real.isOk() ? await Result.tryPromise(() => stat(real.value)) : null;
    if (!real.isOk() || info === null || !info.isOk() || !info.value.isDirectory()) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.usage,
          message: `--root ${root} is not an existing directory.`,
        }),
      );
    }
    resolved.push(real.value);
  }
  return Result.ok(resolved);
};

const isWithin = (root: string, candidate: string): boolean => {
  const relative = path.relative(root, candidate);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
};

/** The real path of `target`, or of its parent joined with its name when it does not exist. */
const realTarget = async (target: string): Promise<string | null> => {
  const real = await Result.tryPromise(() => realpath(target));
  if (real.isOk()) return real.value;
  const parent = await Result.tryPromise(() => realpath(path.dirname(target)));
  return parent.isOk() ? path.join(parent.value, path.basename(target)) : null;
};

type CheckWithinRootsOptions = {
  roots: AllowedRoots;
  /** As the caller passed it; a relative path resolves against the first root. */
  target: string;
  /** Which argument named the path, for the refusal message. */
  argument: string;
};

/**
 * Refuse a path outside every allowed root, returning the resolved path the
 * call should use.
 */
export const checkWithinRoots = async ({
  roots,
  target,
  argument,
}: CheckWithinRootsOptions): Promise<Result<string, FolioCliError>> => {
  const base = roots.at(0) ?? process.cwd();
  const absolute = path.resolve(base, target);
  const real = await realTarget(absolute);
  if (real !== null && roots.some((root) => isWithin(root, real))) {
    return Result.ok(real);
  }
  return Result.err(
    cliError({
      code: FOLIO_CLI_ERROR_CODES.outsideRoot,
      message: `${argument} ${JSON.stringify(target)} is outside the allowed roots.`,
      hint: `Use a path under ${roots.join(", ")}, or restart the server with another --root.`,
    }),
  );
};
