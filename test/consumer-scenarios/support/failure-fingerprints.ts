import { createHash } from "node:crypto";

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
