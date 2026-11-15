/**
 * Which property failures come only from random exploration, and whether a
 * run reports them instead of failing on them.
 *
 * A property's generated pass searches a sequence its seed chooses. When
 * nothing fixed that seed (no registry entry, no caller seed, no
 * `PROPERTY_TEST_SEED`), the sequence is exploration: the commit under CI, or
 * fast-check locally, picked it. A failure there is a new finding, not a
 * regression of a checked-in case. Pinned seeds, caller seeds, replay paths
 * and the property's own `examples` stay regressions everywhere.
 *
 * `FOLIO_FUZZ_EXPLORATION=report` (set by pull request and merge group CI)
 * turns an exploration failure into a report: the FOLIO_FAILURE line and the
 * failure record are written, the test passes. Unset, or outside CI, the
 * failure fails the test as before.
 *
 * A change's own properties are the exception: a test file the change adds
 * or changes (or whose pinned seeds it changes), listed as a JSON array in
 * `FOLIO_FUZZ_EXPLORATION_OWNED` by scripts/ci-test-shards.ts, fails on any
 * failure, so a new property that fails cannot land as a report.
 */

export const EXPLORATION_ENV = "FOLIO_FUZZ_EXPLORATION";
/** Where report-only failures leave their failure records for filing. */
export const EXPLORATION_DIR_ENV = "FOLIO_FUZZ_EXPLORATION_DIR";
/** The commit a report run's change is compared with. */
export const EXPLORATION_BASE_ENV = "FOLIO_FUZZ_EXPLORATION_BASE";
/** JSON array of the repo-relative test files the change under CI owns. */
export const EXPLORATION_OWNED_ENV = "FOLIO_FUZZ_EXPLORATION_OWNED";

/** What chose the seed a fast-check run started from. */
export type SeedSource =
  /** A test/property-seeds/ entry. */
  | "registry"
  /** The property's own `seed` parameter. */
  | "caller"
  /** `PROPERTY_TEST_SEED`. */
  | "environment"
  /** Derived from the commit under CI. */
  | "commit"
  /** None: fast-check drew one. */
  | "fast-check";

export type ExplorationContext = {
  seedSource: SeedSource;
  /** The run started from a counterexample path (its own or `PROPERTY_TEST_PATH`). */
  replayPath: boolean;
  /** How many hand-written `examples` precede the generated values. */
  examples: number;
};

/**
 * Whether a failure at `counterexamplePath` comes only from a randomly
 * generated case. Anything this cannot establish counts as a regression.
 */
export const isExplorationFailure = (
  { seedSource, replayPath, examples }: ExplorationContext,
  counterexamplePath: string | null,
): boolean => {
  if (seedSource !== "commit" && seedSource !== "fast-check") return false;
  if (replayPath) return false;
  if (counterexamplePath === null) return false;
  const head = counterexamplePath.split(":")[0] ?? "";
  if (!/^\d+$/u.test(head)) return false;
  // fast-check yields the examples first; the path starts at the failing
  // value's index in that stream.
  return Number(head) >= examples;
};

const isCi = (env: NodeJS.ProcessEnv): boolean => {
  const raw = env["CI"];
  return raw !== undefined && raw !== "" && raw !== "false" && raw !== "0";
};

/**
 * The path to pin for a counterexample. A pinned replay runs without the
 * property's `examples`, so its path counts generated values only, while the
 * path fast-check reports starts at the failing value's index in a stream that
 * yields the examples first. `null`: the failure is one of the examples, which
 * run on every pass and need no pin, or its index cannot be read.
 */
export const pinnedReplayPath = (
  counterexamplePath: string,
  examples: number,
): string | null => {
  if (examples === 0) return counterexamplePath;
  const [head = "", ...rest] = counterexamplePath.split(":");
  if (!/^\d+$/u.test(head) || Number(head) < examples) return null;
  return [String(Number(head) - examples), ...rest].join(":");
};

/**
 * Whether exploration failures are reported instead of failing the run. Only
 * under CI, where the FOLIO_FAILURE line is always logged, so nothing is
 * dropped silently.
 */
export const explorationReportsOnly = (env: NodeJS.ProcessEnv): boolean =>
  isCi(env) && env[EXPLORATION_ENV] === "report";

/**
 * The test files the change under CI adds or changes, from
 * EXPLORATION_OWNED_ENV. A malformed list throws rather than reporting every
 * failure.
 */
export const explorationOwnedFiles = (
  env: NodeJS.ProcessEnv,
): ReadonlySet<string> => {
  const raw = env[EXPLORATION_OWNED_ENV];
  if (raw === undefined || raw === "") return new Set();
  const parsed: unknown = JSON.parse(raw);
  if (
    !Array.isArray(parsed) ||
    !parsed.every((file): file is string => typeof file === "string")
  )
    throw new Error(
      `${EXPLORATION_OWNED_ENV} must be a JSON array of repo-relative test files`,
    );
  return new Set(parsed);
};

/**
 * Whether an exploration failure in the test file `file` (repo-relative) is
 * reported instead of failing the run: report mode under CI, and the file is
 * not one the change adds or changes.
 */
export const explorationFailureReports = (
  env: NodeJS.ProcessEnv,
  file: string | undefined,
): boolean =>
  explorationReportsOnly(env) &&
  (file === undefined || !explorationOwnedFiles(env).has(file));
