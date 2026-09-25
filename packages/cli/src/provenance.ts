/**
 * Who and when a transaction records. Every revision, comment, and reply a
 * transaction authors carries one author and one UTC timestamp; neither is
 * ever invented.
 */

import { Result } from "better-result";
import { spawnSync } from "node:child_process";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

export const AUTHOR_ENV = "FOLIO_AUTHOR";

type ResolveAuthorOptions = {
  /** `--author`, or the MCP server's configured author. */
  explicit: string | undefined;
  env: Readonly<Record<string, string | undefined>>;
  /** Directory whose git configuration supplies the fallback. */
  cwd: string;
};

const nonEmpty = (value: string | undefined): string | undefined => {
  const trimmed = value?.trim();
  return trimmed === undefined || trimmed === "" ? undefined : trimmed;
};

const gitUserName = (
  cwd: string,
  env: Readonly<Record<string, string | undefined>>,
): string | undefined => {
  const childEnv: Record<string, string> = {};
  for (const [key, value] of Object.entries(env)) {
    if (value !== undefined) childEnv[key] = value;
  }
  const run = spawnSync("git", ["config", "user.name"], {
    cwd,
    env: childEnv,
    encoding: "utf8",
    timeout: 2000,
    stdio: ["ignore", "pipe", "ignore"],
  });
  return run.status === 0 ? nonEmpty(run.stdout) : undefined;
};

/** `--author`, then `FOLIO_AUTHOR`, then git `user.name`; refuses when none is set. */
export const resolveAuthor = ({
  explicit,
  env,
  cwd,
}: ResolveAuthorOptions): Result<string, FolioCliError> => {
  const author = nonEmpty(explicit) ?? nonEmpty(env[AUTHOR_ENV]) ?? gitUserName(cwd, env);
  return author === undefined
    ? Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.authorRequired,
          message: "No author is configured for this change.",
          hint: `Pass --author, set ${AUTHOR_ENV}, or configure git user.name.`,
        }),
      )
    : Result.ok(author);
};

/** An ISO-8601 UTC timestamp at second precision, as revision dates are written. */
const toRevisionDate = (date: Date): string => date.toISOString().replace(/\.\d{3}Z$/u, "Z");

/** The transaction timestamp: `--date` when given (for reproducible output), else now. */
export const resolveTransactionDate = (
  explicit: string | undefined,
  now: () => Date = () => new Date(),
): Result<string, FolioCliError> => {
  if (explicit === undefined) {
    return Result.ok(toRevisionDate(now()));
  }
  const parsed = new Date(explicit);
  return Number.isNaN(parsed.getTime())
    ? Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.usage,
          message: `--date ${JSON.stringify(explicit)} is not an ISO-8601 date.`,
          hint: "Use a form like 2026-01-31T09:00:00Z.",
        }),
      )
    : Result.ok(toRevisionDate(parsed));
};
