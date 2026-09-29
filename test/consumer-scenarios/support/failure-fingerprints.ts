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
};

export const shellQuote = (value: string): string => `'${value.replaceAll("'", `'\\''`)}'`;

const deepestMessage = (failure: unknown): string => {
  let current = failure;
  const seen = new Set<unknown>();
  while (current instanceof Error && current.cause instanceof Error && !seen.has(current.cause)) {
    seen.add(current);
    current = current.cause;
  }
  return current instanceof Error ? current.message : String(current);
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

export const failureMarker = ({
  test,
  seed,
  path,
  repro,
  failure,
}: {
  test: string;
  seed: number;
  path?: string | null;
  repro: string;
  failure: unknown;
}): FailureMarker => {
  const assertion = normalizeAssertion(failure);
  const fingerprint = createHash("sha256")
    .update(`${test}\0${assertion}`)
    .digest("hex")
    .slice(0, 16);
  return { fingerprint, test, seed, path: path ?? null, repro, assertion };
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
  const message = failure instanceof Error ? failure.message : String(failure);
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
