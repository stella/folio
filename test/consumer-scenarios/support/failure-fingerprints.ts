import { stripVTControlCharacters } from "node:util";
import { createHash } from "node:crypto";
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";

export type FailureMarker = {
  fingerprint: string;
  test: string;
  seed: number;
  path: string | null;
  repro: string;
  assertion: string;
  /** Where the compared values differed, values and positions left out (see `diffShape`). */
  diff?: string;
  /**
   * The operations of the minimized flow, step by step, when the fingerprint
   * includes them; `primary` is then the fingerprint without them, the one
   * the failure had before it was shrunk.
   */
  flow?: string;
  primary?: string;
};

export const shellQuote = (value: string): string => `'${value.replaceAll("'", `'\\''`)}'`;

const deepestMessage = (failure: unknown): string => {
  let current = failure;
  const seen = new Set<unknown>();
  while (current instanceof Error && current.cause instanceof Error && !seen.has(current.cause)) {
    seen.add(current);
    current = current.cause;
  }
  return stripVTControlCharacters(current instanceof Error ? current.message : String(current));
};

/** Keep the mismatch kind while removing generated values and document prose. */
export const normalizeAssertion = (failure: unknown): string => {
  const lines = deepestMessage(failure)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const first = lines[0] ?? "unknown failure";
  const symptom = /the result is not what was asked/u.test(first) ? (lines[1] ?? first) : first;
  return symptom
    .replace(/^step\s+-?\d+:\s*/u, "")
    .replace(/\s+\{.*$/u, "")
    .replace(/\s+\[\{.*$/u, "")
    .replace(/"(?:\\.|[^"\\])*"/gu, "<text>")
    .replace(/\b[0-9a-f]{8}-[0-9a-f-]{27,}\b/giu, "<id>")
    .replace(/\b[0-9a-f]{8}\b/giu, "<id>")
    .replace(/-?\d+(?:\.\d+)?/gu, "<n>")
    .replace(/\s+/gu, " ")
    .trim();
};

// A difference line of support/metamorphic.ts `differences`: `path: before → after`,
// perhaps after a relation's own prefix.
const DIFFERENCE =
  /^(?:batch → one at a time )?(\(root\)|(?:\[[^\]\s]*\]|\.[^\s:.[]+)+): (.+?) → (.+)$/u;

/** What kind of value a difference line shows, without the value itself. */
const valueKind = (text: string, path: string): string => {
  const value = text.trim();
  if (value === "undefined") return "∅";
  if (value === "null" || value === "true" || value === "false") return value;
  if (/^-?\d/u.test(value)) return "number";
  if (value.startsWith("{")) {
    // A change, a block or a comment: its type is part of what went wrong.
    const type = /"type":"([\w-]+)"/u.exec(value)?.[1];
    return type === undefined ? "{}" : `{${type}}`;
  }
  if (value.startsWith("[")) return "[]";
  // A `type` field names a kind (insertion, deletion, …), not document text.
  const kind = /\.type$/u.test(path) ? /^…?"([\w-]*)"?/u.exec(value)?.[1] : undefined;
  return kind ?? "text";
};

/**
 * The shape of a comparison's differences: which fields differ, between
 * what kinds of values and which change types, with positions, ids and text
 * left out. Two bugs behind the same symptom line differ here; one bug seen
 * under two seeds does not. Empty when the failure lists no differences.
 */
export const diffShape = (failure: unknown): string => {
  const shapes = new Set<string>();
  for (const line of deepestMessage(failure).split("\n")) {
    const match = DIFFERENCE.exec(line.trim());
    if (match === null) continue;
    const [, at, before, after] = match as unknown as [string, string, string, string];
    // `[2]` → `[]`; a typed entry, `[2:insertion]`, keeps its type.
    const path = at.replace(/\[\d+(?::([\w-]+))?\]/gu, (_, type?: string) =>
      type === undefined ? "[]" : `[${type}]`,
    );
    shapes.add(`${path} ${valueKind(before, path)}→${valueKind(after, path)}`);
  }
  return [...shapes].sort().slice(0, 12).join("; ");
};

const hash = (text: string): string => createHash("sha256").update(text).digest("hex").slice(0, 16);

/**
 * The marker for a failure. The fingerprint hashes the test, the symptom
 * line and the differences' shape; given `flow` (a minimized flow's
 * operations, see `flowShape` in support/fuzz-loop.ts) it hashes that too,
 * and `primary` keeps the fingerprint without it.
 */
export const failureMarker = ({
  test,
  seed,
  path,
  repro,
  failure,
  flow,
}: {
  test: string;
  seed: number;
  path?: string | null;
  repro: string;
  failure: unknown;
  flow?: string;
}): FailureMarker => {
  const assertion = normalizeAssertion(failure);
  const diff = diffShape(failure);
  // Without differences the fingerprint is what it always was.
  const primary = hash(diff === "" ? `${test}\0${assertion}` : `${test}\0${assertion}\0${diff}`);
  return {
    fingerprint: flow === undefined ? primary : hash(`${primary}\0${flow}`),
    test,
    seed,
    path: path ?? null,
    repro,
    assertion,
    ...(diff === "" ? {} : { diff }),
    ...(flow === undefined ? {} : { flow, primary }),
  };
};

export const logFailureMarker = (marker: FailureMarker): void => {
  console.error(`FOLIO_FAILURE ${JSON.stringify(marker)}`);
};

/**
 * What a fuzz run keeps of one failure for the issue that files it: the
 * marker, every one-line replay (most direct first), and, when the flow was
 * shrunk, the minimized flow and how far it shrank.
 */
export type FailureRecord = {
  version: 1;
  marker: FailureMarker;
  replays: string[];
  /** The failure's message, cut to a readable length. */
  error: string;
  flow?: unknown;
  shrink?: { steps: number; from: number; attempts: number };
};

const MAX_RECORD_ERROR = 4_000;

/** A record for `marker`; `error` is the failure itself, cut to length. */
export const failureRecord = (
  marker: FailureMarker,
  failure: unknown,
  extra: Partial<Pick<FailureRecord, "replays" | "flow" | "shrink">> = {},
): FailureRecord => {
  const message = deepestMessage(failure);
  return {
    version: 1,
    marker,
    replays: extra.replays ?? [marker.repro],
    error: message.length > MAX_RECORD_ERROR ? `${message.slice(0, MAX_RECORD_ERROR)} …` : message,
    ...(extra.flow === undefined ? {} : { flow: extra.flow }),
    ...(extra.shrink === undefined ? {} : { shrink: extra.shrink }),
  };
};

/** Write `record` into `dir` as `<fingerprint>-<seed>.json`; returns the path. */
export const writeFailureRecord = (dir: string, record: FailureRecord): string => {
  mkdirSync(dir, { recursive: true });
  const file = join(dir, `${record.marker.fingerprint}-${record.marker.seed}.json`);
  writeFileSync(file, `${JSON.stringify(record, null, 2)}\n`);
  return file;
};

export const reportScenarioFailure = ({
  test,
  seed,
  repro,
  failure,
}: {
  test: string;
  seed: number;
  repro: string;
  failure: unknown;
}): never => {
  logFailureMarker(failureMarker({ test, seed, repro, failure }));
  const message = failure instanceof Error ? failure.message : String(failure);
  throw new Error(`Replay: ${repro}\n${message}`, { cause: failure });
};
