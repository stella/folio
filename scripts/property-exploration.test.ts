import { afterEach, describe, expect, spyOn, test } from "bun:test";
import fc from "fast-check";
/**
 * Exploration reports (test/property-exploration.ts): under pull request and
 * merge group CI a property failure that only the commit's seed found is
 * reported and filed, not failed; every fixed case still fails.
 */
import { mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  type ExplorationContext,
  explorationReportsOnly,
  isExplorationFailure,
} from "../test/property-exploration";
import {
  assertKnownProperty,
  assertProperty,
  overridePinnedSeedsForTesting,
  propertyConfig,
  propertyTestTimeout,
} from "../test/property-testing";
import { extractFailureMarkers } from "./failure-fingerprints";
import { blockingFindings, propertyReplays } from "./fuzz-ejection";
import { parseFailureRecord, readFindings } from "./fuzz-failure-issues";

const FILE = "scripts/property-exploration.test.ts";

const ENV_KEYS = [
  "CI",
  "GITHUB_SHA",
  "PROPERTY_TEST_SEED",
  "PROPERTY_TEST_PATH",
  "PROPERTY_TEST_NUM_RUNS_FACTOR",
  "PROPERTY_TEST_SEED_SALT",
  "FOLIO_FUZZ_HEALTH",
  "FOLIO_FUZZ_EXPLORATION",
  "FOLIO_FUZZ_EXPLORATION_DIR",
  "FOLIO_PROPERTY_FAILURES_DIR",
] as const;
const savedEnv = Object.fromEntries(
  ENV_KEYS.map((key) => [key, process.env[key]]),
);
const directories: string[] = [];

afterEach(() => {
  for (const key of ENV_KEYS) {
    const value = savedEnv[key];
    if (value === undefined) Reflect.deleteProperty(process.env, key);
    else process.env[key] = value;
  }
  overridePinnedSeedsForTesting(undefined);
  for (const directory of directories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

const withEnv = (
  env: Partial<Record<(typeof ENV_KEYS)[number], string>>,
): void => {
  for (const key of ENV_KEYS) Reflect.deleteProperty(process.env, key);
  Object.assign(process.env, env);
};

/** Pull request CI: a commit, report mode and a record directory. */
const reportingCi = (
  extra: Partial<Record<(typeof ENV_KEYS)[number], string>> = {},
): string => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-exploration-"));
  directories.push(directory);
  withEnv({
    CI: "true",
    GITHUB_SHA: "e".repeat(40),
    FOLIO_FUZZ_EXPLORATION: "report",
    FOLIO_FUZZ_EXPLORATION_DIR: directory,
    ...extra,
  });
  return directory;
};

/** Run `body` with console.error captured; returns the logged lines. */
const logged = (body: () => void): string[] => {
  const spy = spyOn(console, "error").mockImplementation(() => {});
  try {
    body();
    return spy.mock.calls.map((parts) => parts.join(" "));
  } finally {
    spy.mockRestore();
  }
};

const generated: ExplorationContext = {
  seedSource: "commit",
  replayPath: false,
  examples: 0,
};

describe("which failures are exploration", () => {
  test("only a commit- or fast-check-chosen seed explores", () => {
    expect(isExplorationFailure(generated, "3:1")).toBe(true);
    expect(
      isExplorationFailure({ ...generated, seedSource: "fast-check" }, "0"),
    ).toBe(true);
    for (const seedSource of ["registry", "caller", "environment"] as const)
      expect(isExplorationFailure({ ...generated, seedSource }, "3:1")).toBe(
        false,
      );
  });

  test("a replay path, an example or an unreadable path is a regression", () => {
    expect(isExplorationFailure({ ...generated, replayPath: true }, "3")).toBe(
      false,
    );
    const withExamples = { ...generated, examples: 2 };
    expect(isExplorationFailure(withExamples, "1:4")).toBe(false);
    expect(isExplorationFailure(withExamples, "2:4")).toBe(true);
    expect(isExplorationFailure(generated, null)).toBe(false);
    expect(isExplorationFailure(generated, "")).toBe(false);
  });

  test("only CI with report mode reports", () => {
    expect(
      explorationReportsOnly({ CI: "true", FOLIO_FUZZ_EXPLORATION: "report" }),
    ).toBe(true);
    expect(explorationReportsOnly({ FOLIO_FUZZ_EXPLORATION: "report" })).toBe(
      false,
    );
    expect(
      explorationReportsOnly({ CI: "true", FOLIO_FUZZ_EXPLORATION: "fail" }),
    ).toBe(false);
    expect(explorationReportsOnly({ CI: "true" })).toBe(false);
  });
});

describe("report mode", () => {
  test(
    "a generated failure is logged, recorded and passes",
    () => {
      const directory = reportingCi();
      const lines = logged(() => {
        assertProperty(
          fc.property(fc.nat(), (value) => value < 10),
          { id: "generated failure reports", numRuns: 200 },
        );
      });
      const markers = extractFailureMarkers(lines.join("\n"));
      expect(markers).toHaveLength(1);
      expect(markers[0]).toMatchObject({
        test: `${FILE}::generated failure reports`,
        exploration: true,
      });
      expect(
        lines.some((line) =>
          line.startsWith("::warning title=Fuzz exploration finding::"),
        ),
      ).toBe(true);
      const property = lines.find((line) =>
        line.startsWith("PROPERTY_FAILURE "),
      );
      expect(
        JSON.parse((property as string).slice("PROPERTY_FAILURE ".length)),
      ).toMatchObject({ exploration: true, pinned: false });
      const records = readdirSync(directory).filter(
        (file) => !file.endsWith(".flow.json"),
      );
      expect(records).toHaveLength(1);
      const record = parseFailureRecord(
        JSON.parse(
          readFileSync(path.join(directory, records[0] as string), "utf8"),
        ),
      );
      expect(record?.marker).toMatchObject({
        fingerprint: markers[0]?.fingerprint,
        exploration: true,
      });
      // The filer reads the record; a merge-group replay leaves it out.
      const findings = readFindings([], [directory]);
      expect(findings).toHaveLength(1);
      expect(blockingFindings(findings)).toEqual([]);
    },
    propertyTestTimeout(5_000),
  );

  test("a caller seed, PROPERTY_TEST_SEED or an example still fails", () => {
    reportingCi();
    logged(() => {
      expect(() => {
        fc.assert(
          fc.property(fc.nat(), (value) => value < 10),
          propertyConfig({ seed: 5, numRuns: 200 }),
        );
      }).toThrow("Replay: ");
    });
    reportingCi({ PROPERTY_TEST_SEED: "5" });
    logged(() => {
      expect(() => {
        assertProperty(
          fc.property(fc.nat(), (value) => value < 10),
          { id: "environment seed fails", numRuns: 200 },
        );
      }).toThrow("Replay: ");
    });
    reportingCi();
    logged(() => {
      expect(() => {
        assertProperty(
          fc.property(fc.nat(), (value) => value !== 1_000_003),
          { id: "example fails", numRuns: 1, examples: [[1_000_003]] },
        );
      }).toThrow("Replay: ");
    });
  });

  test("a pinned seed still fails", () => {
    reportingCi();
    overridePinnedSeedsForTesting({
      [`${FILE}::pinned seed fails`]: [
        { seed: 11, note: "test", date: "2026-10-25" },
      ],
    });
    const lines = logged(() => {
      expect(() => {
        assertProperty(
          fc.property(fc.nat(), () => false),
          { id: "pinned seed fails", numRuns: 5 },
        );
      }).toThrow("Pinned regression seed 11");
    });
    expect(extractFailureMarkers(lines.join("\n"))[0]?.exploration).toBe(
      undefined,
    );
  });

  test("without report mode a generated failure fails", () => {
    withEnv({ CI: "true", GITHUB_SHA: "e".repeat(40) });
    logged(() => {
      expect(() => {
        assertProperty(
          fc.property(fc.nat(), (value) => value < 10),
          { id: "generated failure blocks", numRuns: 200 },
        );
      }).toThrow("Replay: ");
    });
  });

  test("known-failure acceptance keeps failing on a cause that stops reproducing", () => {
    reportingCi();
    const expectedFailures = [
      { family: "never", fingerprint: "0123456789abcdef" },
    ];
    logged(() => {
      expect(() => {
        assertKnownProperty(
          fc.property(fc.nat(), () => true),
          { id: "known cause gone", numRuns: 20, expectedFailures },
        );
      }).toThrow("no longer reproduces");
    });
    // An unrecorded generated failure stops the search: reported, and the
    // causes it had not reached are not judged.
    const lines = logged(() => {
      assertKnownProperty(
        fc.property(fc.nat(), (value) => value < 10),
        { id: "known with a new failure", numRuns: 200, expectedFailures },
      );
    });
    expect(extractFailureMarkers(lines.join("\n"))[0]?.exploration).toBe(true);
  });
});

describe("merge-group replay", () => {
  test("skips exploration reports", () => {
    const report = (exploration: boolean) =>
      `PROPERTY_FAILURE ${JSON.stringify({
        file: "packages/core/src/a.test.ts",
        test: "a",
        seed: 3,
        path: "1",
        numRunsFactor: 1,
        ...(exploration ? { exploration: true } : {}),
      })}`;
    expect(propertyReplays(`${report(true)}\n${report(false)}`)).toEqual([
      {
        file: "packages/core/src/a.test.ts",
        title: "a",
        seed: 3,
        path: "1",
        factor: 1,
      },
    ]);
  });
});
