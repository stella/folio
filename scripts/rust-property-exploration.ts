#!/usr/bin/env bun
/**
 * The exploration pass of a crate's proptest properties.
 *
 * CI first runs `cargo test` with `PROPTEST_CASES=0`: every ordinary test and
 * every checked-in `*.proptest-regressions` seed, blocking. This pass then
 * runs the property targets again with generated cases under a seed derived
 * from the commit (`PROPTEST_RNG_SEED`), so each commit searches afresh and a
 * rerun replays exactly. A test that fails with proptest's
 * `minimal failing input` is a finding: it logs a FOLIO_FAILURE line marked
 * `exploration`, writes a failure record for the merge-group filer and does
 * not fail the run. Any other failure (a panic outside a property, an abort,
 * a crash, a build error) still fails it.
 *
 *   bun scripts/rust-property-exploration.ts --package stella-docx-kernel \
 *     --crate crates/docx-kernel --records test-results/fuzz-exploration
 *
 * `--seed <u32>` replaces the commit seed.
 */

import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { commitSha, hash32 } from "../test/commit-seed";
import {
  type FailureMarker,
  failureMarker,
  failureRecord,
  logFailureMarker,
  shellQuote,
  writeFailureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";

const PROPERTY = /\bproptest!/u;

/** `--lib` when the crate's sources hold properties, and `--test <name>` per integration file that does. */
export const propertyTargets = (
  files: readonly { path: string; text: string }[],
): string[] => {
  const holds = files.filter(({ text }) => PROPERTY.test(text));
  const lib = holds.some(({ path: file }) => file.startsWith("src/"));
  const tests = holds
    .map(({ path: file }) => /^tests\/([^/]+)\.rs$/u.exec(file)?.[1])
    .filter((name): name is string => name !== undefined)
    .toSorted();
  return [
    ...(lib ? ["--lib"] : []),
    ...tests.flatMap((name) => ["--test", name]),
  ];
};

/** The commit's exploration seed: an unsigned 32-bit value proptest reads as u64. */
export const explorationSeed = (sha: string): number =>
  hash32(`${sha}\0rust proptest exploration`) >>> 0;

export type TestBinary = {
  /** `--lib` or `--test <name>`, as cargo takes it. */
  target: string[];
  /** The source the binary was built from, `src/lib.rs` or `tests/<name>.rs`. */
  source: string;
  result: "ok" | "failed" | "missing";
  /** Failed test names, from the binary's `failures:` list. */
  failed: string[];
  /** Each failed test's captured output. */
  outputs: Map<string, string>;
};

const RUNNING = /^\s+Running (?:unittests )?(\S+) \(/u;
const RESULT = /^test result: (ok|FAILED)\./u;
const SECTION = /^---- (\S+) stdout ----$/u;

/** Split `cargo test` output (stdout and stderr merged) into its test binaries. */
export const parseCargoTestOutput = (output: string): TestBinary[] => {
  const binaries: TestBinary[] = [];
  let current: TestBinary | undefined;
  let section: { name: string; lines: string[] } | undefined;
  // libtest prints `failures:` twice: before the outputs, then before the names.
  let headings = 0;
  const close = () => {
    if (section !== undefined && current !== undefined)
      current.outputs.set(section.name, section.lines.join("\n"));
    section = undefined;
  };
  for (const line of output.split("\n")) {
    const running = RUNNING.exec(line);
    if (running !== null) {
      close();
      headings = 0;
      const source = running[1] as string;
      const test = /^tests\/([^/]+)\.rs$/u.exec(source)?.[1];
      current = {
        target: test === undefined ? ["--lib"] : ["--test", test],
        source,
        result: "missing",
        failed: [],
        outputs: new Map(),
      };
      binaries.push(current);
      continue;
    }
    if (current === undefined) continue;
    const header = SECTION.exec(line);
    if (header !== null) {
      close();
      section = { name: header[1] as string, lines: [] };
      continue;
    }
    if (line === "failures:") {
      close();
      headings += 1;
      continue;
    }
    const result = RESULT.exec(line);
    if (result !== null) {
      close();
      headings = 0;
      current.result = result[1] === "ok" ? "ok" : "failed";
      continue;
    }
    if (section !== undefined) {
      section.lines.push(line);
      continue;
    }
    const name = /^ {4}(\S+)$/u.exec(line)?.[1];
    if (headings === 2 && name !== undefined) current.failed.push(name);
  }
  close();
  return binaries;
};

export type ExplorationFinding = {
  target: string[];
  source: string;
  test: string;
  /** proptest's reason, the text between `Test failed: ` and `minimal failing input`. */
  reason: string;
  /** The `cc <seed> # …` line that pins the shrunk case, when proptest printed one. */
  pin: string | null;
};

const FAILED = /Test failed: ([\s\S]*?)\.\nminimal failing input: /u;

/** A failed test's finding, or null when its failure is not a generated case's. */
export const explorationFinding = (
  binary: TestBinary,
  test: string,
): ExplorationFinding | null => {
  const output = binary.outputs.get(test);
  if (output === undefined) return null;
  const failed = FAILED.exec(output);
  if (failed === null) return null;
  return {
    target: binary.target,
    source: binary.source,
    test,
    reason: (failed[1] as string).trim(),
    pin: /^cc [0-9a-f]+\b.*$/mu.exec(output)?.[0] ?? null,
  };
};

/** The findings of a run, and every failure that must still fail it. */
export const classifyRun = (
  binaries: readonly TestBinary[],
  exitCode: number,
): { findings: ExplorationFinding[]; blocking: string[] } => {
  const findings: ExplorationFinding[] = [];
  const blocking: string[] = [];
  for (const binary of binaries) {
    if (binary.result === "missing") {
      blocking.push(`${binary.source}: no test result (crashed or aborted)`);
      continue;
    }
    if (binary.result === "ok") continue;
    if (binary.failed.length === 0)
      blocking.push(`${binary.source}: failed without naming a test`);
    for (const test of binary.failed) {
      const finding = explorationFinding(binary, test);
      if (finding === null)
        blocking.push(
          `${binary.source}: ${test} failed outside a generated case`,
        );
      else findings.push(finding);
    }
  }
  if (exitCode !== 0 && findings.length === 0 && blocking.length === 0)
    blocking.push(
      `cargo test exited ${String(exitCode)} without a failing test`,
    );
  if (binaries.length === 0) blocking.push("cargo test ran no test binary");
  return { findings, blocking };
};

/** The FOLIO_FAILURE marker of a finding. */
export const findingMarker = (
  finding: ExplorationFinding,
  options: { package: string; crate: string; seed: number },
): FailureMarker => ({
  ...failureMarker({
    test: `${options.crate}/${finding.source}::${finding.test}`,
    seed: options.seed,
    repro: `PROPTEST_RNG_SEED=${String(options.seed)} cargo test -p ${options.package} ${finding.target.join(" ")} -- ${shellQuote(finding.test)} --exact`,
    failure: new Error(finding.reason),
  }),
  exploration: true,
});

type Options = {
  package: string;
  crate: string;
  records: string | null;
  seed: number | null;
};

const parseArgs = (argv: readonly string[]): Options => {
  const options: Options = {
    package: "",
    crate: "",
    records: null,
    seed: null,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? "";
    if (arg === "--package") options.package = value();
    else if (arg === "--crate") options.crate = value();
    else if (arg === "--records") options.records = value() || null;
    else if (arg === "--seed") {
      const seed = Number(value());
      if (!Number.isSafeInteger(seed) || seed < 0)
        throw new Error("--seed must be a non-negative integer");
      options.seed = seed;
    } else throw new Error(`Unknown argument ${String(arg)}`);
  }
  if (options.package === "" || options.crate === "")
    throw new Error("--package and --crate are required");
  return options;
};

const crateFiles = (crate: string) =>
  readdirSync(crate, { recursive: true })
    .map(String)
    .filter(
      (file) =>
        file.endsWith(".rs") &&
        (file.startsWith("src/") || /^tests\/[^/]+\.rs$/u.test(file)),
    )
    .map((file) => ({
      path: file,
      text: readFileSync(path.join(crate, file), "utf8"),
    }));

const main = async (): Promise<void> => {
  const options = parseArgs(process.argv.slice(2));
  const sha = commitSha();
  const seed = options.seed ?? (sha === null ? null : explorationSeed(sha));
  if (seed === null)
    throw new Error(
      "No commit to derive the exploration seed from (give --seed)",
    );
  const targets = propertyTargets(crateFiles(options.crate));
  if (targets.length === 0) {
    console.log(`No proptest properties in ${options.crate}.`);
    return;
  }
  console.log(
    `→ proptest exploration of ${options.package} (PROPTEST_RNG_SEED=${String(seed)})`,
  );
  const child = Bun.spawn(
    [
      "sh",
      "-c",
      'exec "$@" 2>&1',
      "sh",
      "cargo",
      "test",
      "-p",
      options.package,
      "--no-fail-fast",
      ...targets,
    ],
    {
      stdout: "pipe",
      env: { ...process.env, PROPTEST_RNG_SEED: String(seed) },
    },
  );
  let output = "";
  const decoder = new TextDecoder();
  for await (const chunk of child.stdout) {
    process.stdout.write(chunk);
    output += decoder.decode(chunk, { stream: true });
  }
  const exitCode = await child.exited;
  const { findings, blocking } = classifyRun(
    parseCargoTestOutput(output),
    exitCode,
  );
  for (const finding of findings) {
    const marker = findingMarker(finding, { ...options, seed });
    logFailureMarker(marker);
    if (options.records !== null)
      writeFailureRecord(
        options.records,
        failureRecord(
          marker,
          new Error(
            finding.pin === null
              ? finding.reason
              : `${finding.reason}\n\nPin: add \`${finding.pin}\` to ${options.crate}/${finding.source.replace(/\.rs$/u, ".proptest-regressions")}`,
          ),
        ),
      );
    console.error(
      `::warning title=Fuzz exploration finding::${marker.test} failed on a generated case (reported, not failing this run). Replay: ${marker.repro}`,
    );
  }
  if (blocking.length > 0) {
    for (const reason of blocking) console.error(`::error::${reason}`);
    process.exit(1);
  }
};

if (import.meta.main) await main();
