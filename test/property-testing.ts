import fc from "fast-check";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { classifyFuzzRun } from "./fuzz-health";
import { reportFuzzHealth } from "./consumer-scenarios/support/fuzz-health";

import { commitSeed } from "./commit-seed";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  writeFailureRecord,
} from "./consumer-scenarios/support/failure-fingerprints";

/**
 * Shared fast-check configuration for the repo's property tests. Every
 * `fc.assert` goes through `propertyConfig` (or `assertProperty`), which
 * centralizes:
 *
 *  1. Longer nightly runtime. `PROPERTY_TEST_NUM_RUNS_FACTOR` scales every
 *     property's `numRuns` (the nightly job sets 10; PR CI raises it for the
 *     property files a change touches, see scripts/property-areas.ts). Unset,
 *     the factor is 1.
 *
 *  2. One seed per commit. Under CI a property with no seed of its own runs
 *     under a seed derived from the commit (`GITHUB_SHA`, else `git rev-parse
 *     HEAD`), its own call site and the factor, so each property searches its
 *     own sequence, every commit searches afresh, and a rerun of the same
 *     commit replays exactly: a red run stays red instead of passing on rerun.
 *     `PROPERTY_TEST_SEED_SALT` (the nightly sets its run id) mixes in more.
 *     Locally fast-check picks a random seed.
 *
 *  3. Replayable failures. A failing property's message ends with a replay
 *     line, `cd <package> && PROPERTY_TEST_SEED=… PROPERTY_TEST_PATH=… bun
 *     test <file> -t '<test>'`, and under CI a `PROPERTY_FAILURE {json}` line
 *     that scripts/nightly-failure-issues.ts turns into an issue. Verbose
 *     reporting is on under CI so the log lists every shrunk failing value.
 *     `PROPERTY_TEST_SEED` pins the seed for every property in the run (a seed
 *     sweep iterates it); `PROPERTY_TEST_PATH` jumps straight to a
 *     counterexample of the property running under that seed. A property that
 *     pins its own `seed` keeps it.
 *
 *  4. Pinned regression seeds. test/property-seeds.json maps
 *     `<repo-relative file>::<test title as written>` to seeds (with an
 *     optional counterexample path) that once failed. `assertProperty`
 *     replays each of them before the property's generated runs, in every
 *     environment; `propertyConfig` refuses a property that has pinned seeds
 *     but is not asserted through `assertProperty`.
 */

const NUM_RUNS_FACTOR_ENV = "PROPERTY_TEST_NUM_RUNS_FACTOR";
const SEED_ENV = "PROPERTY_TEST_SEED";
const PATH_ENV = "PROPERTY_TEST_PATH";
const SEED_SALT_ENV = "PROPERTY_TEST_SEED_SALT";

/** fast-check's own default when a property does not specify `numRuns`. */
const FAST_CHECK_DEFAULT_NUM_RUNS = 100;

/** Longest counterexample carried in the machine-readable failure line. */
const MAX_REPORTED_COUNTEREXAMPLE = 4_000;

const SELF = fileURLToPath(import.meta.url);
const REPO_ROOT = path.resolve(path.dirname(SELF), "..");
export const PROPERTY_SEEDS_FILE = "test/property-seeds.json";

const readNumRunsFactor = (raw: string | undefined): number => {
  if (raw === undefined) {
    return 1;
  }
  const parsed = Number(raw);
  // A factor below 1 (or non-numeric) would silently weaken nightly coverage;
  // fall back to the neutral factor instead.
  return Number.isFinite(parsed) && parsed >= 1 ? parsed : 1;
};

const numRunsFactor = (): number => readNumRunsFactor(process.env[NUM_RUNS_FACTOR_ENV]);

const readSeed = (raw: string | undefined): number | undefined => {
  if (raw === undefined || raw === "") {
    return undefined;
  }
  const parsed = Number(raw);
  // An unreadable seed must not silently turn into fast-check's own random
  // one: a sweep would then report a "failing seed" nothing can replay.
  if (!Number.isInteger(parsed)) {
    throw new Error(`${SEED_ENV} must be an integer, received ${JSON.stringify(raw)}.`);
  }
  return parsed;
};

/**
 * The seed every property in this process runs under, or `undefined` when
 * fast-check should pick its own.
 */
export const propertyTestSeed = (): number | undefined => readSeed(process.env[SEED_ENV]);

// Treat the common CI values as enabled, but honor an explicit opt-out
// (`CI=false`/`0`) so verbose reporting can be silenced locally.
const isCi = (): boolean => {
  const raw = process.env["CI"];
  return raw !== undefined && raw !== "" && raw !== "false" && raw !== "0";
};

// ---------------------------------------------------------------------------
// Call sites and test titles
// ---------------------------------------------------------------------------

type CallSite = {
  /** Absolute path of the file that called in. */
  absolute: string;
  /** Repo-relative path, `/`-separated. */
  file: string;
  line: number;
};

const FRAME = /\(?((?:file:\/\/)?(?:\/|[A-Za-z]:[\\/])[^()]*?):(\d+):\d+\)?\s*$/;

const parseFrame = (frame: string): { absolute: string; line: number } | undefined => {
  const match = FRAME.exec(frame);
  if (match === null) {
    return undefined;
  }
  const raw = match[1] as string;
  return {
    absolute: raw.startsWith("file://") ? fileURLToPath(raw) : raw,
    line: Number(match[2]),
  };
};

/**
 * The test-file line that called into this module: the first frame outside it
 * that is a `*.test.ts(x)` file, else the first frame outside it.
 */
const callSite = (): CallSite | undefined => {
  const limit = Error.stackTraceLimit;
  Error.stackTraceLimit = 64;
  const stack = new Error("call site").stack ?? "";
  Error.stackTraceLimit = limit;
  const frames = stack
    .split("\n")
    .slice(1)
    .map(parseFrame)
    .filter((frame) => frame !== undefined && frame.absolute !== SELF);
  const frame = frames.find((candidate) => /\.test\.tsx?$/.test(candidate!.absolute)) ?? frames[0];
  if (frame === undefined) {
    return undefined;
  }
  return {
    ...frame,
    file: path.relative(REPO_ROOT, frame.absolute).replaceAll("\\", "/"),
  };
};

const sourceLines = new Map<string, readonly string[]>();

const linesOf = (absolute: string): readonly string[] => {
  let lines = sourceLines.get(absolute);
  if (lines === undefined) {
    try {
      lines = readFileSync(absolute, "utf8").split("\n");
    } catch {
      lines = [];
    }
    sourceLines.set(absolute, lines);
  }
  return lines;
};

const indentOf = (text: string): number => text.length - text.trimStart().length;

const OPENER = /^\s*(?:describe|test|it)\b[\w.]*(?:\([^()]*\))?\(\s*(.*)$/;
const TITLE = /^(["'`])((?:\\.|(?!\1)[^\\])*)\1/;

/** The title a `describe`/`test`/`it` opener at `index` gives, as written. */
const openerTitle = (lines: readonly string[], index: number): string | undefined => {
  const opener = OPENER.exec(lines[index] ?? "");
  if (opener === null) {
    return undefined;
  }
  let rest = (opener[1] as string).trim();
  for (let next = index + 1; rest === "" && next < lines.length; next += 1) {
    rest = (lines[next] as string).trim();
  }
  return TITLE.exec(rest)?.[2];
};

/**
 * The `describe`/`test` titles enclosing `line` (1-based), outermost first,
 * read from the formatted source by indentation. Titles are as written, so a
 * template title keeps its `${…}` placeholders.
 */
export const enclosingTitles = (lines: readonly string[], line: number): string[] => {
  let indent = indentOf(lines[line - 1] ?? "");
  const titles: string[] = [];
  for (let index = line - 2; index >= 0 && indent > 0; index -= 1) {
    const text = lines[index] as string;
    if (text.trim() === "" || indentOf(text) >= indent || /^\s*[)}\]]/.test(text)) {
      continue;
    }
    indent = indentOf(text);
    const title = openerTitle(lines, index);
    if (title !== undefined) {
      titles.unshift(title);
    }
  }
  return titles;
};

type PropertyIdentity = {
  site: CallSite | undefined;
  /** The enclosing test's title as written, when one could be read. */
  title: string | undefined;
  /** `<file>::<title>`, the property-seeds.json key. */
  key: string | undefined;
};

const identify = (): PropertyIdentity => {
  const site = callSite();
  if (site === undefined) {
    return { site, title: undefined, key: undefined };
  }
  const title = enclosingTitles(linesOf(site.absolute), site.line).at(-1);
  return { site, title, key: title === undefined ? undefined : `${site.file}::${title}` };
};

// ---------------------------------------------------------------------------
// Pinned regression seeds
// ---------------------------------------------------------------------------

export type ExpectedPropertyFailure = {
  family: string;
  fingerprint: string;
};

export type KnownPropertyFailure<Ts> = ExpectedPropertyFailure & {
  matches?: (value: Ts) => boolean;
};

export type PinnedSeed = {
  seed: number;
  path?: string;
  note: string;
  date: string;
  expectedFailure?: ExpectedPropertyFailure;
};

let pinnedSeeds: Record<string, readonly PinnedSeed[]> | undefined;

/** test/property-seeds.json, keyed `<file>::<test title>`; `$`-keys are comments. */
export const readPinnedSeeds = (): Record<string, readonly PinnedSeed[]> => {
  if (pinnedSeeds === undefined) {
    const parsed = JSON.parse(
      readFileSync(path.join(REPO_ROOT, PROPERTY_SEEDS_FILE), "utf8"),
    ) as Record<string, unknown>;
    pinnedSeeds = Object.fromEntries(
      Object.entries(parsed).filter(
        (entry): entry is [string, PinnedSeed[]] => !entry[0].startsWith("$"),
      ),
    );
  }
  return pinnedSeeds;
};

let pinnedSeedsOverride: Record<string, readonly PinnedSeed[]> | undefined;

/** Stand in for test/property-seeds.json (this module's own tests); `undefined` restores it. */
export const overridePinnedSeedsForTesting = (
  seeds: Record<string, readonly PinnedSeed[]> | undefined,
): void => {
  pinnedSeedsOverride = seeds;
};

const pinnedFor = (key: string | undefined): readonly PinnedSeed[] =>
  key === undefined ? [] : ((pinnedSeedsOverride ?? readPinnedSeeds())[key] ?? []);

// ---------------------------------------------------------------------------
// Replay reporting
// ---------------------------------------------------------------------------

const packageDirOf = (absolute: string): string => {
  let dir = path.dirname(absolute);
  while (dir.startsWith(REPO_ROOT) && dir !== REPO_ROOT) {
    try {
      readFileSync(path.join(dir, "package.json"));
      return dir;
    } catch {
      dir = path.dirname(dir);
    }
  }
  return REPO_ROOT;
};

const shellQuote = (text: string): string => `'${text.replaceAll("'", `'\\''`)}'`;

/** A `bun test -t` pattern for a title as written: `${…}` matches anything. */
export const titlePattern = (title: string): string =>
  title
    .split(/\$\{[^}]*\}/)
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join(".*");

type Replay = { seed: number; path: string | null };

const replayCommand = (identity: PropertyIdentity, replay: Replay): string => {
  const env = [`${SEED_ENV}=${String(replay.seed)}`];
  if (replay.path !== null && replay.path !== "") {
    env.push(`${PATH_ENV}=${shellQuote(replay.path)}`);
  }
  const factor = numRunsFactor();
  if (factor !== 1) {
    env.push(`${NUM_RUNS_FACTOR_ENV}=${String(factor)}`);
  }
  const { site, title } = identity;
  if (site === undefined) {
    return `${env.join(" ")} bun test`;
  }
  const packageDir = packageDirOf(site.absolute);
  const file = path.relative(packageDir, site.absolute).replaceAll("\\", "/");
  const cd =
    packageDir === REPO_ROOT
      ? ""
      : `cd ${path.relative(REPO_ROOT, packageDir).replaceAll("\\", "/")} && `;
  const filter = title === undefined ? "" : ` -t ${shellQuote(titlePattern(title))}`;
  return `${cd}${env.join(" ")} bun test ./${file}${filter}`;
};

const truncate = (text: string, max: number): string =>
  text.length > max
    ? `${text.slice(0, max)}… (${String(text.length - max)} more characters)`
    : text;

const errorText = (error: unknown): string => {
  if (error instanceof Error) {
    return error.message;
  }
  return typeof error === "string" ? error : fc.stringify(error);
};

/**
 * A reporter that throws what fast-check would, plus the replay line and, for
 * a property with no pinned entry for this seed, the property-seeds.json entry
 * that would pin it. Under CI it also logs one machine-readable line.
 */
const replayReporter =
  <Ts>(identity: PropertyIdentity, pinned: PinnedSeed | undefined) =>
  (details: fc.RunDetails<Ts>): void => {
    const health = classifyFuzzRun(details);
    reportFuzzHealth(health);
    if (health.status === "infrastructure") {
      throw new Error(`Fuzz infrastructure: ${health.detail}`);
    }
    const expected = pinned?.expectedFailure;
    if (!details.failed) {
      if (expected !== undefined) {
        throw new Error(
          `${expected.family} no longer reproduces: pinned seed ${String(pinned?.seed)} passes; remove expectedFailure from ${PROPERTY_SEEDS_FILE}.`,
        );
      }
      return;
    }
    if (expected !== undefined) {
      const marker = failureMarker({
        test: identity.key ?? identity.title ?? "property",
        seed: details.seed,
        path: details.counterexamplePath,
        repro: replayCommand(identity, { seed: details.seed, path: details.counterexamplePath }),
        failure: details.errorInstance,
      });
      if (marker.fingerprint === expected.fingerprint) return;
      throw details.errorInstance;
    }
    const counterexamplePath = details.counterexamplePath;
    const replay = replayCommand(identity, { seed: details.seed, path: counterexamplePath });
    const entry = {
      seed: details.seed,
      ...(counterexamplePath === null ? {} : { path: counterexamplePath }),
    };
    const lines = [fc.defaultReportMessage(details) ?? "Property failed", ""];
    if (pinned !== undefined) {
      lines.push(
        `Pinned regression seed ${String(pinned.seed)} from ${PROPERTY_SEEDS_FILE} failed again (${pinned.note}).`,
      );
    }
    lines.push(`Replay: ${replay}`);
    if (pinned === undefined && identity.key !== undefined) {
      lines.push(
        `Pin: add ${JSON.stringify(entry)} (with a note and date) under ${JSON.stringify(identity.key)} in ${PROPERTY_SEEDS_FILE}`,
      );
    }
    if (isCi()) {
      const directory = process.env["FOLIO_PROPERTY_FAILURES_DIR"];
      const marker = failureMarker({
        test: identity.key ?? identity.title ?? "property",
        seed: details.seed,
        path: counterexamplePath,
        repro: replay,
        failure: details.errorInstance,
      });
      logFailureMarker(marker);
      if (directory !== undefined)
        writeFailureRecord(
          directory,
          failureRecord(marker, details.errorInstance, { flow: details.counterexample }),
        );
      console.error(
        `PROPERTY_FAILURE ${JSON.stringify({
          file: identity.site?.file ?? null,
          line: identity.site?.line ?? null,
          test: identity.title ?? null,
          key: identity.key ?? null,
          seed: details.seed,
          path: counterexamplePath,
          numRunsFactor: numRunsFactor(),
          pinned: pinned !== undefined,
          replay,
          counterexample: truncate(
            fc.stringify(details.counterexample),
            MAX_REPORTED_COUNTEREXAMPLE,
          ),
          error: truncate(errorText(details.errorInstance), 1_000),
        })}`,
      );
    }
    throw new Error(lines.join("\n"), { cause: details.errorInstance });
  };

// ---------------------------------------------------------------------------
// Seeds and parameters
// ---------------------------------------------------------------------------

/** The seed a property with no seed of its own runs under, or `undefined` to let fast-check pick. */
const defaultSeed = (identity: PropertyIdentity): number | undefined => {
  const explicit = propertyTestSeed();
  if (explicit !== undefined) {
    return explicit;
  }
  if (!isCi()) {
    return undefined;
  }
  const salt = [
    identity.site === undefined ? "" : `${identity.site.file}:${String(identity.site.line)}`,
    String(numRunsFactor()),
    process.env[SEED_SALT_ENV] ?? "",
  ].join("\0");
  return commitSeed(salt);
};

/**
 * `PROPERTY_TEST_PATH` for a property running under the `PROPERTY_TEST_SEED`
 * it was found with; otherwise the property's own `path`, if any.
 */
const envReplayPath = (seed: number | undefined, own: string | undefined): string | undefined => {
  const envPath = process.env[PATH_ENV];
  const applies =
    envPath !== undefined && envPath !== "" && seed !== undefined && seed === propertyTestSeed();
  return applies ? envPath : own;
};

const buildConfig = <Ts>(
  params: fc.Parameters<Ts>,
  identity: PropertyIdentity,
  pinned?: PinnedSeed,
): fc.Parameters<Ts> => {
  const baseNumRuns = params.numRuns ?? FAST_CHECK_DEFAULT_NUM_RUNS;
  const seed = pinned?.seed ?? params.seed ?? defaultSeed(identity);
  const replayPath = pinned === undefined ? envReplayPath(seed, params.path) : pinned.path;
  const own = params.reporter === undefined && params.asyncReporter === undefined;
  const expected = pinned?.expectedFailure;
  const config: fc.Parameters<Ts> = {
    verbose: isCi(),
    ...params,
    ...(seed === undefined ? {} : { seed }),
    ...(replayPath === undefined ? {} : { path: replayPath }),
    ...(own ? { reporter: replayReporter<Ts>(identity, pinned) } : {}),
    numRuns: Math.ceil(baseNumRuns * numRunsFactor()),
  };
  if (expected !== undefined) {
    // A recorded counterexample is one exact replay, independent of the nightly
    // factor and caller reporters. The generated pass keeps its own configuration.
    Reflect.deleteProperty(config, "asyncReporter");
    config.reporter = replayReporter<Ts>(identity, pinned);
    config.numRuns = 1;
    config.endOnFailure = true;
  }
  return config;
};

/**
 * Build the `fc.assert` parameters for a property test: pass the per-test
 * tuning you want in PR CI (typically just `numRuns`) and this scales it for
 * the nightly sweep, seeds it per commit under CI, enables verbose reporting
 * under CI and makes a failure print its replay line.
 *
 * ```ts
 * fc.assert(fc.property(arb, predicate), propertyConfig({ numRuns: 200 }));
 * ```
 */
export const propertyConfig = <Ts>(params: fc.Parameters<Ts> = {}): fc.Parameters<Ts> =>
  configFor(params, identify(), false);

const configFor = <Ts>(
  params: fc.Parameters<Ts>,
  identity: PropertyIdentity,
  viaAssertProperty: boolean,
): fc.Parameters<Ts> => {
  if (!viaAssertProperty && pinnedFor(identity.key).length > 0) {
    throw new Error(
      `${String(identity.key)} has pinned seeds in ${PROPERTY_SEEDS_FILE}, which only replay through ` +
        "assertProperty: write `assertProperty(property, { numRuns })` instead of " +
        "`fc.assert(property, propertyConfig({ numRuns }))`.",
    );
  }
  return buildConfig(params, identity);
};

/**
 * `fc.assert(property, propertyConfig(params))`, after first replaying every
 * seed test/property-seeds.json pins for the calling test.
 */
export function assertProperty<Ts>(
  property: fc.IAsyncProperty<Ts>,
  params?: fc.Parameters<Ts>,
): Promise<void>;
export function assertProperty<Ts>(property: fc.IProperty<Ts>, params?: fc.Parameters<Ts>): void;
export function assertProperty<Ts>(
  property: fc.IRawProperty<Ts>,
  params: fc.Parameters<Ts> = {},
): Promise<void> | void {
  const identity = identify();
  const pinned = pinnedFor(identity.key);
  // Examples run in the generated pass; including them in a pinned replay
  // shifts fast-check's path indices away from the recorded counterexample.
  const replays = pinned.map((entry) => buildConfig({ ...params, examples: [] }, identity, entry));
  const generated = configFor(params, identity, true);
  return runConfiguredProperty(property, [...replays, generated]);
}

/** Run the raw law for pinned seeds and require every recorded cause in generated cases. */
export function assertKnownProperty<Ts>(
  property: fc.IAsyncProperty<Ts>,
  expectedFailures: readonly KnownPropertyFailure<Ts>[],
  params?: fc.Parameters<Ts>,
): Promise<void>;
export function assertKnownProperty<Ts>(
  property: fc.IProperty<Ts>,
  expectedFailures: readonly KnownPropertyFailure<Ts>[],
  params?: fc.Parameters<Ts>,
): void;
export function assertKnownProperty<Ts>(
  property: fc.IRawProperty<Ts>,
  expectedFailures: readonly KnownPropertyFailure<Ts>[],
  params: fc.Parameters<Ts> = {},
): Promise<void> | void {
  if (expectedFailures.length === 0) throw new Error("Known property requires a recorded cause");
  const identity = identify();
  const seen = new Set<KnownPropertyFailure<Ts>>();
  const inspect = (value: Ts, outcome: fc.PreconditionFailure | fc.PropertyFailure | null) => {
    if (outcome === null || outcome instanceof fc.PreconditionFailure) return outcome;
    const marker = failureMarker({
      test: identity.key ?? identity.title ?? "property",
      seed: 0,
      repro: "",
      failure: outcome.error,
    });
    const matched = expectedFailures.filter(
      (expected) =>
        expected.fingerprint === marker.fingerprint &&
        (expected.matches === undefined || expected.matches(value)),
    );
    if (matched.length === 0) return outcome;
    for (const expected of matched) seen.add(expected);
    return null;
  };
  const search: fc.IRawProperty<Ts> = {
    isAsync: () => property.isAsync(),
    generate: (mrng, runId) => property.generate(mrng, runId),
    shrink: (value) => property.shrink(value),
    runBeforeEach: () => property.runBeforeEach(),
    runAfterEach: () => property.runAfterEach(),
    run: (value) => {
      const result = property.run(value);
      return result instanceof Promise
        ? result.then((outcome) => inspect(value, outcome))
        : inspect(value, result);
    },
  };
  const generated = configFor(params, identity, true);
  Reflect.deleteProperty(generated, "asyncReporter");
  generated.reporter = (details) => {
    replayReporter<Ts>(identity, undefined)(details);
    const missing = expectedFailures.filter((expected) => !seen.has(expected));
    if (missing.length > 0) {
      throw new Error(
        `${missing.map(({ family, fingerprint }) => `${family}:${fingerprint}`).join(", ")} no longer reproduces: fixed: set this kind to holds.`,
      );
    }
  };
  const replays = pinnedFor(identity.key).map((entry) =>
    buildConfig({ ...params, examples: [] }, identity, entry),
  );
  if (property.isAsync()) {
    return (async () => {
      await runConfiguredProperty(property, replays);
      await runConfiguredProperty(search, [generated]);
    })();
  }
  runConfiguredProperty(property, replays);
  return runConfiguredProperty(search, [generated]);
}

/** Run only registry seeds for bugs already fixed, without a fresh generated pass. */
export function assertPinnedProperty<Ts>(
  property: fc.IAsyncProperty<Ts>,
  params?: fc.Parameters<Ts>,
): Promise<void>;
export function assertPinnedProperty<Ts>(
  property: fc.IProperty<Ts>,
  params?: fc.Parameters<Ts>,
): void;
export function assertPinnedProperty<Ts>(
  property: fc.IRawProperty<Ts>,
  params: fc.Parameters<Ts> = {},
): Promise<void> | void {
  const identity = identify();
  const pinned = pinnedFor(identity.key);
  if (pinned.length === 0)
    throw new Error(`No fixed regression seeds registered for ${String(identity.key)}`);
  const replays = pinned.map((entry) => {
    const config = buildConfig({ ...params, examples: [] }, identity, entry);
    config.numRuns = 1;
    return config;
  });
  return runConfiguredProperty(property, replays);
}

const runConfiguredProperty = <Ts>(
  property: fc.IRawProperty<Ts>,
  configs: readonly fc.Parameters<Ts>[],
): Promise<void> | void => {
  if (property.isAsync()) {
    return (async () => {
      for (const config of configs) {
        reportFuzzHealth({ status: "started", completed: 0 });
        await fc.assert(property, config);
      }
    })();
  }
  for (const config of configs) {
    reportFuzzHealth({ status: "started", completed: 0 });
    fc.assert(property, config);
  }
};

/**
 * Scale a per-test Bun timeout (ms) by the same nightly factor that scales
 * `numRuns`, so an expensive property whose run count grows ×N also gets ×N
 * wall-clock before it is killed. In PR CI (factor 1) the timeout is unchanged.
 *
 * ```ts
 * test("round-trip", () => { ... }, propertyTestTimeout(15_000));
 * ```
 */
export const propertyTestTimeout = (baseMs: number): number => Math.ceil(baseMs * numRunsFactor());
