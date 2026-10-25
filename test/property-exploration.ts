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
 */

export const EXPLORATION_ENV = "FOLIO_FUZZ_EXPLORATION";
/** Where report-only failures leave their failure records for filing. */
export const EXPLORATION_DIR_ENV = "FOLIO_FUZZ_EXPLORATION_DIR";

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
 * Whether exploration failures are reported instead of failing the run. Only
 * under CI, where the FOLIO_FAILURE line is always logged, so nothing is
 * dropped silently.
 */
export const explorationReportsOnly = (env: NodeJS.ProcessEnv): boolean =>
  isCi(env) && env[EXPLORATION_ENV] === "report";
