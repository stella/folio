/**
 * The replay line a failing property prints, as data and as text.
 *
 * test/property-testing.ts writes the line; scripts that replay a reported
 * failure on another commit read it back with `parseReplayCommand`. Both go
 * through this module, so the reader cannot drift from the writer.
 */
import fc from "fast-check";

import { hash32 } from "./commit-seed";
import { NUM_RUNS_FACTOR_ENV } from "./property-run-factor";

export const SEED_ENV = "PROPERTY_TEST_SEED";
export const PATH_ENV = "PROPERTY_TEST_PATH";
/**
 * The digest of the counterexample a replay line's path leads to. A replay
 * whose path leads elsewhere (the generator changed) fails as stale instead
 * of reporting a pass for a case that was never the recorded one.
 */
export const CASE_ENV = "PROPERTY_TEST_CASE";

/** One property replay. */
export type ReplayCommand = {
  /** The package directory, relative to the repository root; `""` for the root. */
  dir: string;
  /** The test file relative to `dir`; absent when the property has no call site. */
  file?: string;
  /** The `bun test -t` pattern, when the line filters by test name. */
  pattern?: string;
  seed: number;
  path: string | null;
  /** The digest of the counterexample `path` leads to (`counterexampleDigest`). */
  digest?: string;
  factor: number;
};

/** A digest of a generated value, as the replay line carries it. */
export const counterexampleDigest = (value: unknown): string =>
  (hash32(fc.stringify(value)) >>> 0).toString(16).padStart(8, "0");

const shellQuote = (text: string): string =>
  `'${text.replaceAll("'", `'\\''`)}'`;

/** The replay line for `command`. */
export const formatReplayCommand = (command: ReplayCommand): string => {
  const env = [`${SEED_ENV}=${String(command.seed)}`];
  if (command.path !== null && command.path !== "") {
    env.push(`${PATH_ENV}=${shellQuote(command.path)}`);
    if (command.digest !== undefined) env.push(`${CASE_ENV}=${command.digest}`);
  }
  if (command.factor !== 1)
    env.push(`${NUM_RUNS_FACTOR_ENV}=${String(command.factor)}`);
  if (command.file === undefined) return `${env.join(" ")} bun test`;
  const cd = command.dir === "" ? "" : `cd ${command.dir} && `;
  const filter =
    command.pattern === undefined ? "" : ` -t ${shellQuote(command.pattern)}`;
  return `${cd}${env.join(" ")} bun test ./${command.file}${filter}`;
};

const QUOTED = String.raw`'((?:[^']|'\\'')*)'`;
const REPLAY_LINE = new RegExp(
  [
    String.raw`^(?:cd ([^\s']+) && )?`,
    `${SEED_ENV}=(-?\\d+)`,
    `(?: ${PATH_ENV}=${QUOTED})?`,
    `(?: ${CASE_ENV}=([0-9a-f]{8}))?`,
    `(?: ${NUM_RUNS_FACTOR_ENV}=(\\d+(?:\\.\\d+)?))?`,
    String.raw` bun test(?: \./([^\s']+))?`,
    `(?: -t ${QUOTED})?$`,
  ].join(""),
  "u",
);

const unquote = (text: string): string => text.replaceAll(`'\\''`, "'");

/** `line` as a replay, or `undefined` when it is not a replay line. */
export const parseReplayCommand = (line: string): ReplayCommand | undefined => {
  const match = REPLAY_LINE.exec(line);
  if (match === null) return undefined;
  const [, dir, seed, path, digest, factor, file, pattern] = match;
  const parsedSeed = Number(seed);
  if (!Number.isSafeInteger(parsedSeed)) return undefined;
  return {
    dir: dir ?? "",
    ...(file === undefined ? {} : { file }),
    ...(pattern === undefined ? {} : { pattern: unquote(pattern) }),
    seed: parsedSeed,
    path: path === undefined ? null : unquote(path),
    ...(digest === undefined ? {} : { digest }),
    factor: factor === undefined ? 1 : Number(factor),
  };
};
