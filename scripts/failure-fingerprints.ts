import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";

import type { FailureMarker } from "../test/consumer-scenarios/support/failure-fingerprints";

type KnownFailure = {
  fingerprint: string;
  issueOrPr: string;
  firstSeen: string;
};

const MARKER = "FOLIO_FAILURE ";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

export const parseKnownFailures = (value: unknown): KnownFailure[] => {
  if (!isRecord(value) || !Array.isArray(value["known"])) {
    throw new Error("Invalid known failure fingerprint registry");
  }
  return value["known"].map((entry: unknown) => {
    if (
      !isRecord(entry) ||
      typeof entry["fingerprint"] !== "string" ||
      !/^[0-9a-f]{16}$/u.test(entry["fingerprint"]) ||
      typeof entry["issueOrPr"] !== "string" ||
      entry["issueOrPr"].length === 0 ||
      typeof entry["firstSeen"] !== "string" ||
      !/^\d{4}-\d{2}-\d{2}$/u.test(entry["firstSeen"])
    ) {
      throw new Error("Invalid known failure fingerprint entry");
    }
    return {
      fingerprint: entry["fingerprint"],
      issueOrPr: entry["issueOrPr"],
      firstSeen: entry["firstSeen"],
    };
  });
};

export const extractFailureMarkers = (log: string): FailureMarker[] =>
  log.split("\n").flatMap((line) => {
    const start = line.indexOf(MARKER);
    if (start === -1) return [];
    const value: unknown = JSON.parse(line.slice(start + MARKER.length).trim());
    if (
      typeof value !== "object" ||
      value === null ||
      !("fingerprint" in value) ||
      !("test" in value) ||
      !("seed" in value) ||
      !("repro" in value) ||
      !("path" in value) ||
      !("assertion" in value) ||
      typeof value.fingerprint !== "string" ||
      typeof value.test !== "string" ||
      typeof value.seed !== "number" ||
      typeof value.repro !== "string" ||
      (value.path !== null && typeof value.path !== "string") ||
      typeof value.assertion !== "string"
    ) {
      throw new Error("Malformed FOLIO_FAILURE line");
    }
    const optional = value as { diff?: unknown; flow?: unknown; primary?: unknown };
    return [
      {
        fingerprint: value.fingerprint,
        test: value.test,
        seed: value.seed,
        repro: value.repro,
        path: value.path,
        assertion: value.assertion,
        ...(typeof optional.diff === "string" ? { diff: optional.diff } : {}),
        ...(typeof optional.flow === "string" && typeof optional.primary === "string"
          ? { flow: optional.flow, primary: optional.primary }
          : {}),
      },
    ];
  });

export const classifyFailureMarkers = (
  markers: readonly FailureMarker[],
  known: readonly KnownFailure[],
) => {
  const knownByFingerprint = new Map(known.map((entry) => [entry.fingerprint, entry]));
  if (knownByFingerprint.size !== known.length) {
    throw new Error("Duplicate known failure fingerprint");
  }
  const grouped = new Map<
    string,
    { marker: FailureMarker; seeds: Set<number>; known: KnownFailure | undefined }
  >();
  for (const marker of markers) {
    const entry = grouped.get(marker.fingerprint);
    if (entry) {
      entry.seeds.add(marker.seed);
      continue;
    }
    grouped.set(marker.fingerprint, {
      marker,
      seeds: new Set([marker.seed]),
      known: knownByFingerprint.get(marker.fingerprint),
    });
  }
  return [...grouped.values()].map(({ marker, seeds, known: entry }) => {
    const result = {
      status: entry === undefined ? "new" : "known",
      fingerprint: marker.fingerprint,
      test: marker.test,
      seeds: [...seeds],
      repro: marker.repro,
    };
    if (entry !== undefined) {
      Object.assign(result, { issueOrPr: entry.issueOrPr, firstSeen: entry.firstSeen });
    }
    return result;
  });
};

const usage = "Usage: bun scripts/failure-fingerprints.ts <run-id|job-log-file>";

if (import.meta.main) {
  const source = process.argv[2];
  if (source === "--help" || source === "-h") {
    console.log(usage);
  } else if (source === undefined || process.argv.length !== 3) {
    throw new Error(usage);
  } else {
    const log = /^\d+$/u.test(source)
      ? execFileSync("gh", ["run", "view", source, "--log"], {
          encoding: "utf8",
          maxBuffer: 50 * 1024 * 1024,
        })
      : readFileSync(path.resolve(source), "utf8");
    const registry: unknown = JSON.parse(
      readFileSync(new URL("../test/known-failure-fingerprints.json", import.meta.url), "utf8"),
    );
    for (const entry of classifyFailureMarkers(
      extractFailureMarkers(log),
      parseKnownFailures(registry),
    )) {
      console.log(JSON.stringify(entry));
    }
  }
}
