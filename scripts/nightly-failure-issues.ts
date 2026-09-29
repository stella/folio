#!/usr/bin/env bun
/**
 * Turn a failed nightly test log into GitHub issues: one per failing property
 * test, one per (operation, placement, violation kind) for the conformance
 * tier, so a class of failure that hits many document shapes files once.
 *
 * The nightly property sweep and the full conformance tier run unattended, so
 * a failure nobody reads is a failure nobody fixes. This parses the run's log
 * for each failing test (and, for a property, its seed, counterexample path
 * and counterexample) and opens an issue labelled `nightly-property-failure`
 * (or `nightly-conformance-failure`), or comments on the open one for the same
 * test. A property issue carries the line that replays the failure locally
 * and the exact test/property-seeds.json entry that pins it once it is fixed;
 * seeds are never committed automatically.
 *
 * It reads the `PROPERTY_FAILURE {json}` lines test/property-testing.ts logs
 * under CI, and falls back to fast-check's own report (`{ seed: …, path: … }`,
 * `Counterexample: …`) for a property that did not go through it.
 *
 * Usage:
 *   bun scripts/nightly-failure-issues.ts --kind property|conformance
 *     --log <file> [--run-url <url>] [--sha <sha>] [--factor <n>] [--dry-run]
 *
 * `--dry-run` prints the issues instead of calling `gh`.
 */

import { $ } from "bun";
import { createHash } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

export type Kind = "property" | "conformance";

export type Failure = {
  /** The full test name bun printed, `describe > test`. */
  name: string;
  /** Repo-relative test file, when the log names it. */
  file: string | null;
  seed: number | null;
  path: string | null;
  counterexample: string | null;
  /** The test title as written in the source (a template keeps `${…}`). */
  title: string | null;
  replay: string | null;
  error: string | null;
  /** A seed test/property-seeds.json already pins failed again. */
  pinned: boolean;
  /** Conformance violation kinds the error names, `mode kind` (e.g. `editing undo`). */
  violations?: string[];
  /** A conformance group: every shape where one operation at one placement broke one way. */
  group?: ConformanceGroup;
};

export type ConformanceGroup = {
  operation: string;
  placement: string;
  kind: string;
  shapes: string[];
  modes: string[];
};

const MAX_ISSUES = 20;
const MAX_TEXT = 6_000;
const LABELS: Record<Kind, { name: string; color: string; description: string }> = {
  property: {
    name: "nightly-property-failure",
    color: "d93f0b",
    description: "A property the nightly sweep found failing",
  },
  conformance: {
    name: "nightly-conformance-failure",
    color: "d93f0b",
    description: "A test the nightly full conformance tier found failing",
  },
};

// `gh run view --log` prefixes `job\tstep\t<timestamp> `; `bun --filter`
// prefixes `<package> <script>: `.
const GH_LOG_PREFIX = /^[^\t]*\t[^\t]*\t\d{4}-\d\d-\d\dT[\d:.]+Z ?/;
const FILTER_PREFIX = /^@?[\w.-]+(?:\/[\w.-]+)? test(?::[\w-]+)*: /;
const FAIL_LINE = /^\(fail\) (.+?)(?: \[[\d.]+m?s\])?$/;
const GROUP_FILE = /^(?:::group::)?((?:[\w.@-]+\/)*[\w.@-]+\.test\.tsx?):$/;
const SEED_LINE = /^\{ seed: (-?\d+), path: "([\d:]*)", endOnFailure: true \}$/;
const ERROR_FILE = /::error file=((?:[\w.@-]+\/)*[\w.@-]+\.test\.tsx?),/;
const FRAME_FILE = /\/(packages\/[\w.-]+\/(?:[\w.@-]+\/)*[\w.@-]+\.test\.tsx?):\d+:\d+\)?$/;
/** A received violation in a conformance case's diff: `+   "[editing] undo: …",`. */
const VIOLATION_LINE = /^\+\s+"\[([\w-]+)\] ([\w-]+):/;
/** A conformance case: `… > shape › operation @ placement`. */
const CASE_NAME = /(?:^| > )([^>]+?) › (.+) @ ([^@]+)$/;
export const CONFORMANCE_FILE = "packages/core/src/__tests__/editorCommandConformance.test.ts";
const CONFORMANCE_COMMAND =
  "cd packages/core && FOLIO_CONFORMANCE=full bun test src/__tests__/editorCommandConformance.test.ts";

type Marker = {
  file: string | null;
  test: string | null;
  seed: number;
  path: string | null;
  replay: string;
  counterexample: string;
  error: string;
  pinned: boolean;
};

const clean = (line: string): string =>
  line.replace(GH_LOG_PREFIX, "").replace(FILTER_PREFIX, "").replace(/\r$/, "");

/** Every failing test in a bun test log, with the property details it carries. */
export const parseFailures = (
  log: string,
  packageOf?: (file: string) => string | null,
): Failure[] => {
  const lines = log.split("\n").map(clean);
  const failures = new Map<string, Failure>();
  let groupFile: string | null = null;
  let located: string | null = null;
  let pending: Partial<Failure> = {};
  let violations = new Set<string>();
  let marker: Marker | null = null;
  let summary = false;

  const resolveGroup = (file: string): string => {
    if (file.startsWith("packages/") || file.startsWith("scripts/")) return file;
    const pkg = packageOf?.(file) ?? null;
    return pkg === null ? file : `${pkg}/${file}`;
  };

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index] as string;
    const group = GROUP_FILE.exec(line.trim());
    if (group !== null) {
      groupFile = resolveGroup(group[1] as string);
      summary = false;
      continue;
    }
    if (line.trim() === "::endgroup::") {
      groupFile = null;
      continue;
    }
    if (/^\d+ tests? failed:$/.test(line.trim())) {
      summary = true;
      continue;
    }
    if (summary) continue;
    if (line.startsWith("PROPERTY_FAILURE ")) {
      try {
        marker = JSON.parse(line.slice("PROPERTY_FAILURE ".length)) as Marker;
      } catch {
        marker = null;
      }
      continue;
    }
    const seed = SEED_LINE.exec(line.trim());
    if (seed !== null) {
      pending.seed = Number(seed[1]);
      pending.path = seed[2] === "" ? null : (seed[2] as string);
      continue;
    }
    if (line.startsWith("Counterexample: ") && pending.counterexample === undefined) {
      pending.counterexample = line.slice("Counterexample: ".length);
      continue;
    }
    if (/^error: /.test(line) && !line.startsWith("error: Property failed") && !pending.error) {
      const detail = [line.slice("error: ".length)];
      for (let next = index + 1; next < lines.length && next < index + 8; next += 1) {
        const text = lines[next] as string;
        if (/^\s+at /.test(text) || FAIL_LINE.test(text)) break;
        detail.push(text);
      }
      pending.error = detail.join("\n").trim();
    }
    const violation = VIOLATION_LINE.exec(line.trim());
    if (violation !== null) {
      violations.add(`${violation[1] as string} ${violation[2] as string}`);
    }
    const errorFile = ERROR_FILE.exec(line) ?? FRAME_FILE.exec(line);
    if (errorFile !== null && located === null) {
      located = errorFile[1] as string;
    }
    const fail = FAIL_LINE.exec(line);
    if (fail === null) continue;
    const name = fail[1] as string;
    const file = marker?.file ?? located ?? groupFile;
    const key = `${file ?? "<unknown file>"}::${name}`;
    if (!failures.has(key)) {
      failures.set(key, {
        name,
        file,
        seed: marker?.seed ?? pending.seed ?? null,
        path: marker === null ? (pending.path ?? null) : marker.path,
        counterexample: marker?.counterexample ?? pending.counterexample ?? null,
        title: marker?.test ?? null,
        replay: marker?.replay ?? null,
        error: marker?.error ?? pending.error ?? null,
        pinned: marker?.pinned ?? false,
        ...(violations.size === 0 ? {} : { violations: [...violations] }),
      });
    }
    pending = {};
    violations = new Set();
    marker = null;
    located = null;
  }
  return [...failures.values()];
};

const shellQuote = (text: string): string => `'${text.replaceAll("'", `'\\''`)}'`;
const escapeRegExp = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** The innermost test title of a `describe > test` name. */
const lastSegment = (name: string): string => name.split(" > ").at(-1) as string;

const packageDirOf = (file: string): string | null =>
  /^(packages\/[^/]+)\//.exec(file)?.[1] ?? null;

/** The command that reruns a whole nightly suite: a replay for a failure no single test names. */
export const suiteReplay = (kind: Kind): string =>
  kind === "conformance" ? CONFORMANCE_COMMAND : "bun run test:property";

/** The `FOLIO_CONFORMANCE_FILTER` regex that selects exactly these cases. */
export const caseFilter = (shapes: readonly string[], operation: string, placement: string) => {
  const shape =
    shapes.length === 1
      ? escapeRegExp(shapes[0] as string)
      : `(?:${shapes.map(escapeRegExp).join("|")})`;
  return `^${shape} › ${escapeRegExp(operation)} @ ${escapeRegExp(placement)}$`;
};

/** A replay command for a failure whose log did not carry one. */
export const replayFor = (kind: Kind, failure: Failure, factor: number | null): string => {
  if (failure.replay !== null) return failure.replay;
  const pattern = shellQuote(escapeRegExp(lastSegment(failure.name)));
  if (kind === "conformance") {
    if (failure.group !== undefined) {
      const { shapes, operation, placement } = failure.group;
      const filter = shellQuote(caseFilter(shapes, operation, placement));
      // A replacer function: the filter ends in `$'`, a special pattern in a replacement string.
      return CONFORMANCE_COMMAND.replace(
        "FOLIO_CONFORMANCE=full ",
        () => `FOLIO_CONFORMANCE=full FOLIO_CONFORMANCE_FILTER=${filter} `,
      );
    }
    // A case (`shape › operation @ placement`) replays alone under the filter.
    const test = lastSegment(failure.name);
    const filter = / › .+ @ /.test(test)
      ? `FOLIO_CONFORMANCE_FILTER=${shellQuote(`^${escapeRegExp(test)}$`)} `
      : "";
    return `cd packages/core && FOLIO_CONFORMANCE=full ${filter}bun test src/__tests__/editorCommandConformance.test.ts -t ${pattern}`;
  }
  if (failure.file === null) return suiteReplay(kind);
  const env: string[] = [];
  if (failure.seed !== null) env.push(`PROPERTY_TEST_SEED=${String(failure.seed)}`);
  if (failure.path !== null) env.push(`PROPERTY_TEST_PATH=${shellQuote(failure.path)}`);
  if (factor !== null && factor !== 1) env.push(`PROPERTY_TEST_NUM_RUNS_FACTOR=${String(factor)}`);
  const file = failure.file;
  const pkg = packageDirOf(file);
  const cd = pkg === null ? "" : `cd ${pkg} && `;
  const relative = pkg === null ? file : file.slice(pkg.length + 1);
  return `${cd}${[...env, "bun test", relative].join(" ")} -t ${pattern}`;
};

/**
 * Conformance failures grouped by (operation, placement, violation kind): a
 * case that broke two ways joins two groups, and a group lists every shape
 * and mode it broke in. A failure that is not a case (a coverage test, the
 * stale-gap check) stays on its own.
 */
export const groupConformanceFailures = (failures: readonly Failure[]): Failure[] => {
  const groups = new Map<string, Failure & { group: ConformanceGroup }>();
  const rest: Failure[] = [];
  for (const failure of failures) {
    const match = CASE_NAME.exec(failure.name);
    if (match === null) {
      rest.push(failure);
      continue;
    }
    const [, shape, operation, placement] = match as unknown as [string, string, string, string];
    const byKind = new Map<string, Set<string>>();
    for (const violation of failure.violations ?? []) {
      const [mode, kind] = violation.split(" ") as [string, string];
      byKind.set(kind, (byKind.get(kind) ?? new Set()).add(mode));
    }
    if (byKind.size === 0) byKind.set("unclassified", new Set());
    for (const [kind, modes] of byKind) {
      const name = `${operation} @ ${placement}: ${kind}`;
      let group = groups.get(name);
      if (group === undefined) {
        group = {
          ...failure,
          name,
          file: failure.file ?? CONFORMANCE_FILE,
          group: { operation, placement, kind, shapes: [], modes: [] },
        };
        delete group.violations;
        groups.set(name, group);
      }
      if (!group.group.shapes.includes(shape)) group.group.shapes.push(shape);
      for (const mode of modes) {
        if (!group.group.modes.includes(mode)) group.group.modes.push(mode);
      }
    }
  }
  for (const group of groups.values()) {
    group.group.modes.sort();
    group.replay = replayFor("conformance", { ...group, replay: null }, null);
  }
  return [...groups.values(), ...rest];
};

/**
 * One issue for the failures past the first `MAX_ISSUES`, under a title that
 * stays the same from run to run, listing each with its own replay line.
 */
export const overflowFailure = (
  kind: Kind,
  rest: readonly Failure[],
  factor: number | null,
): Failure => ({
  ...unparsedFailure(kind),
  name: `more failing ${kind === "conformance" ? "groups" : "tests"} than one run files`,
  error: rest.map((failure) => `${failure.name}\n  ${replayFor(kind, failure, factor)}`).join("\n"),
});

/** The test/property-seeds.json entry that pins a property failure. */
export const seedEntry = (failure: Failure, date: string, runUrl: string | null): string | null => {
  if (failure.seed === null || failure.file === null) return null;
  const key = `${failure.file}::${failure.title ?? lastSegment(failure.name)}`;
  const entry = {
    seed: failure.seed,
    ...(failure.path === null ? {} : { path: failure.path }),
    note: `nightly ${date}${runUrl === null ? "" : ` (${runUrl})`}: <what it caught>`,
    date,
  };
  return `${JSON.stringify(key)}: [\n  ${JSON.stringify(entry)}\n]`;
};

const truncate = (text: string): string =>
  text.length > MAX_TEXT
    ? `${text.slice(0, MAX_TEXT)}\n… (${String(text.length - MAX_TEXT)} more characters)`
    : text;

const fence = (text: string, lang = ""): string => {
  const ticks = text.includes("```") ? "````" : "```";
  return `${ticks}${lang}\n${truncate(text)}\n${ticks}`;
};

type Context = {
  kind: Kind;
  runUrl: string | null;
  sha: string | null;
  factor: number | null;
  date: string;
};

export const issueTitle = (kind: Kind, failure: Failure): string => {
  const where = failure.group !== undefined || failure.file === null ? "" : `${failure.file}::`;
  const title = `Nightly ${kind} failure: ${where}${failure.name}`;
  if (title.length <= 240) return title;
  const key = createHash("sha256").update(title).digest("hex").slice(0, 12);
  return `${title.slice(0, 226)}…${key}`;
};

/** The body of a new issue, or of a comment on the open one, for a failure. */
export const issueBody = (failure: Failure, context: Context, recurrence: boolean): string => {
  const where = [
    context.runUrl === null ? "the nightly run" : `[the nightly run](${context.runUrl})`,
    context.sha === null ? null : `commit \`${context.sha.slice(0, 12)}\``,
    context.kind === "property" && context.factor !== null
      ? `factor ${String(context.factor)}`
      : null,
  ]
    .filter((part) => part !== null)
    .join(", ");
  const sections = [
    recurrence
      ? `Failed again in ${where}.`
      : `${failure.seed === null ? "A test" : "A property"} failed in ${where}.`,
    ...(failure.pinned
      ? ["", "This is a seed `test/property-seeds.json` already pins: a fixed failure came back."]
      : []),
    "",
    ...(failure.group === undefined
      ? [`**Test:** ${failure.name}`]
      : [
          `**Operation:** \`${failure.group.operation}\` at \`${failure.group.placement}\``,
          `**Violation:** \`${failure.group.kind}\`${failure.group.modes.length === 0 ? "" : ` (${failure.group.modes.join(", ")})`}`,
          `**Shapes:** ${failure.group.shapes.map((shape) => `\`${shape}\``).join(", ")}`,
        ]),
    ...(failure.file === null ? [] : [`**File:** \`${failure.file}\``]),
    ...(failure.seed === null
      ? []
      : [
          `**Seed:** \`${String(failure.seed)}\`${failure.path === null ? "" : `, **path:** \`${failure.path}\``}`,
        ]),
    "",
    "### Replay",
    fence(replayFor(context.kind, failure, context.factor), "sh"),
  ];
  if (failure.counterexample !== null) {
    sections.push("", "### Counterexample", fence(failure.counterexample));
  }
  if (failure.error !== null && failure.error.trim() !== "") {
    sections.push("", "### Error", fence(failure.error.trim()));
  }
  const entry =
    context.kind === "property" && !failure.pinned
      ? seedEntry(failure, context.date, context.runUrl)
      : null;
  if (entry !== null) {
    sections.push(
      "",
      "### Pin it",
      "Once fixed, add this entry to `test/property-seeds.json` (the test must assert through `assertProperty`, which replays it first in every run):",
      fence(entry, "json"),
    );
  }
  return sections.join("\n");
};

/** One issue for a failed run whose log names no failing test (a crash or timeout). */
export const unparsedFailure = (kind: Kind): Failure => ({
  name: `${kind === "property" ? "property sweep" : "conformance tier"} failed without a failing test in the log`,
  file: null,
  seed: null,
  path: null,
  counterexample: null,
  title: null,
  replay: suiteReplay(kind),
  error: "The job failed before or outside a test (setup, crash or timeout); see the run log.",
  pinned: false,
});

const parseArgs = (argv: readonly string[]) => {
  const options = {
    kind: "property" as Kind,
    log: "",
    runUrl: null as string | null,
    sha: null as string | null,
    factor: null as number | null,
    dryRun: false,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? "";
    if (arg === "--kind") options.kind = value() === "conformance" ? "conformance" : "property";
    else if (arg === "--log") options.log = value();
    else if (arg === "--run-url") options.runUrl = value() || null;
    else if (arg === "--sha") options.sha = value() || null;
    else if (arg === "--factor") options.factor = Number(value()) || null;
    else if (arg === "--dry-run") options.dryRun = true;
    else throw new Error(`Unknown argument ${String(arg)}`);
  }
  return options;
};

/** `src/x.test.ts` in a `bun --filter` log is relative to its package: find which. */
const packageOfFile =
  (root: string) =>
  (file: string): string | null => {
    for (const pkg of [
      "packages/core",
      "packages/docx-core",
      "packages/agents",
      "packages/react",
    ]) {
      if (existsSync(path.join(root, pkg, file))) return pkg;
    }
    return null;
  };

const main = async (): Promise<void> => {
  const options = parseArgs(process.argv.slice(2));
  const root = path.resolve(import.meta.dir, "..");
  const log =
    options.log !== "" && existsSync(options.log) ? readFileSync(options.log, "utf8") : "";
  const found = parseFailures(log, packageOfFile(root));
  const parsed = options.kind === "conformance" ? groupConformanceFailures(found) : found;
  const failures = parsed.length === 0 ? [unparsedFailure(options.kind)] : parsed;
  const context: Context = {
    kind: options.kind,
    runUrl: options.runUrl,
    sha: options.sha,
    factor: options.factor,
    date: new Date().toISOString().slice(0, 10),
  };
  const label = LABELS[options.kind];
  const shown = failures.slice(0, MAX_ISSUES);
  if (failures.length > MAX_ISSUES) {
    shown.push(overflowFailure(options.kind, failures.slice(MAX_ISSUES), options.factor));
  }

  if (options.dryRun) {
    for (const failure of shown) {
      console.log(
        `=== ${issueTitle(options.kind, failure)}\n${issueBody(failure, context, false)}\n`,
      );
    }
    return;
  }

  await $`gh label create ${label.name} --color ${label.color} --description ${label.description} --force`
    .quiet()
    .nothrow();
  const issueEndpoint = `repos/{owner}/{repo}/issues?state=open&labels=${label.name}&per_page=100`;
  const openPages = JSON.parse(await $`gh api --paginate --slurp ${issueEndpoint}`.text()) as {
    number: number;
    title: string;
  }[][];
  const open = openPages.flat();
  for (const failure of shown) {
    const title = issueTitle(options.kind, failure);
    const existing = open.find((issue) => issue.title === title);
    const bodyFile = path.join(tmpdir(), `nightly-failure-${String(process.pid)}.md`);
    writeFileSync(bodyFile, issueBody(failure, context, existing !== undefined));
    if (existing === undefined) {
      const url = (
        await $`gh issue create --title ${title} --label ${label.name} --body-file ${bodyFile}`.text()
      ).trim();
      console.log(`opened ${url}: ${title}`);
    } else {
      await $`gh issue comment ${String(existing.number)} --body-file ${bodyFile}`.quiet();
      console.log(`commented on #${String(existing.number)}: ${title}`);
    }
  }
};

if (import.meta.main) {
  await main();
}
