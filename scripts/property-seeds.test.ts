/**
 * The property-test seed discipline (test/property-testing.ts): per-commit
 * seeds under CI, a replay line on every failure, and the pinned regression
 * seeds in test/property-seeds.json replayed first.
 */

import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { readFileSync } from "node:fs";
import path from "node:path";

import { commitSeed, hash32 } from "../test/commit-seed";
import {
  assertPinnedProperty,
  assertProperty,
  enclosingTitles,
  overridePinnedSeedsForTesting,
  PROPERTY_SEEDS_FILE,
  propertyConfig,
  propertyTestTimeout,
  readPinnedSeeds,
  titlePattern,
} from "../test/property-testing";

setDefaultTimeout(propertyTestTimeout(5_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const ENV_KEYS = [
  "CI",
  "GITHUB_SHA",
  "PROPERTY_TEST_SEED",
  "PROPERTY_TEST_PATH",
  "PROPERTY_TEST_NUM_RUNS_FACTOR",
  "PROPERTY_TEST_SEED_SALT",
] as const;
const savedEnv = Object.fromEntries(ENV_KEYS.map((key) => [key, process.env[key]]));

afterEach(() => {
  for (const key of ENV_KEYS) {
    const value = savedEnv[key];
    if (value === undefined) Reflect.deleteProperty(process.env, key);
    else process.env[key] = value;
  }
});

const withEnv = (env: Partial<Record<(typeof ENV_KEYS)[number], string>>): void => {
  for (const key of ENV_KEYS) Reflect.deleteProperty(process.env, key);
  Object.assign(process.env, env);
};

/** Two sites on different lines, so each gets its own seed. */
const seedAtFirstSite = (): number | undefined => propertyConfig().seed;
const seedAtSecondSite = (): number | undefined => propertyConfig().seed;

describe("per-commit seeds", () => {
  test("under CI a property's seed follows the commit, its call site and the factor", () => {
    withEnv({ CI: "true", GITHUB_SHA: "a".repeat(40) });
    const first = seedAtFirstSite();
    expect(first).toBeNumber();
    expect(seedAtFirstSite()).toBe(first);
    expect(seedAtSecondSite()).not.toBe(first);
    expect(commitSeed("salt")).toBe(commitSeed("salt"));

    process.env["PROPERTY_TEST_NUM_RUNS_FACTOR"] = "5";
    expect(seedAtFirstSite()).not.toBe(first);
  });

  test("locally fast-check picks the seed; PROPERTY_TEST_SEED pins it everywhere", () => {
    withEnv({});
    expect(seedAtFirstSite()).toBeUndefined();
    withEnv({ CI: "true", PROPERTY_TEST_SEED: "42" });
    expect(seedAtFirstSite()).toBe(42);
    expect(propertyConfig({ seed: 7 }).seed).toBe(7);
  });

  test("the hash spreads one-character salts apart", () => {
    const seeds = new Set(Array.from({ length: 1000 }, (_, index) => hash32(`s${String(index)}`)));
    expect(seeds.size).toBe(1000);
    expect([...seeds].every((seed) => (seed | 0) === seed)).toBe(true);
  });

  test("PROPERTY_TEST_PATH applies only under the seed it was found with", () => {
    withEnv({ PROPERTY_TEST_SEED: "9", PROPERTY_TEST_PATH: "3:1" });
    expect(propertyConfig().path).toBe("3:1");
    expect(propertyConfig({ seed: 8 }).path).toBeUndefined();
  });
});

describe("failure reporting", () => {
  test("a failure names the seed, the path and a replay line for this test", () => {
    withEnv({});
    let message = "";
    try {
      fc.assert(
        fc.property(fc.nat(), (value) => value < 50),
        propertyConfig({ seed: 1234, numRuns: 200 }),
      );
    } catch (error) {
      message = (error as Error).message;
    }
    expect(message).toContain("seed: 1234");
    const replay = message.split("\n").find((line) => line.startsWith("Replay: "));
    expect(replay).toMatch(
      /^Replay: PROPERTY_TEST_SEED=1234 PROPERTY_TEST_PATH='[\d:]+' bun test scripts\/property-seeds\.test\.ts -t 'a failure names the seed, the path and a replay line for this test'$/,
    );
    expect(message).toContain(
      `under "scripts/property-seeds.test.ts::a failure names the seed, the path and a replay line for this test" in ${PROPERTY_SEEDS_FILE}`,
    );
  });

  test("the replay line replays the failure", () => {
    withEnv({});
    const failing = (): string => {
      try {
        fc.assert(
          fc.property(fc.array(fc.nat()), (values) => values.length < 3),
          propertyConfig({ numRuns: 200 }),
        );
      } catch (error) {
        return (error as Error).message;
      }
      return "";
    };
    const first = failing();
    const seed = /PROPERTY_TEST_SEED=(-?\d+)/.exec(first)?.[1];
    const shrunkPath = /PROPERTY_TEST_PATH='([\d:]+)'/.exec(first)?.[1];
    expect(seed).toBeDefined();
    withEnv({ PROPERTY_TEST_SEED: seed as string, PROPERTY_TEST_PATH: shrunkPath as string });
    const replayed = failing();
    expect(/Counterexample: (.*)/.exec(replayed)?.[1]).toBe(
      /Counterexample: (.*)/.exec(first)?.[1] as string,
    );
  });

  test("under CI a failure also logs one machine-readable line", () => {
    withEnv({ CI: "true", GITHUB_SHA: "c".repeat(40) });
    const logged: string[] = [];
    const original = console.error;
    console.error = (...parts: unknown[]) => {
      logged.push(parts.join(" "));
    };
    try {
      expect(() => {
        fc.assert(
          fc.property(fc.nat(), (value) => value < 10),
          propertyConfig({ numRuns: 200 }),
        );
      }).toThrow();
    } finally {
      console.error = original;
    }
    const line = logged.find((entry) => entry.startsWith("PROPERTY_FAILURE "));
    const report = JSON.parse((line as string).slice("PROPERTY_FAILURE ".length)) as Record<
      string,
      unknown
    >;
    expect(report).toMatchObject({
      file: "scripts/property-seeds.test.ts",
      test: "under CI a failure also logs one machine-readable line",
      counterexample: "[10]",
      pinned: false,
    });
    expect(report["seed"]).toBeNumber();
    const markerLines = logged.filter((entry) => entry.startsWith("FOLIO_FAILURE "));
    expect(markerLines).toHaveLength(1);
    const markerLine = markerLines[0];
    const marker = JSON.parse((markerLine as string).slice("FOLIO_FAILURE ".length)) as Record<
      string,
      unknown
    >;
    expect(marker).toMatchObject({
      test: "scripts/property-seeds.test.ts::under CI a failure also logs one machine-readable line",
      assertion: "Property failed by returning false",
    });
    expect(marker["seed"]).toBe(report["seed"]);
    expect(marker["repro"]).toBe(report["replay"]);
  });
});

describe("pinned regression seeds", () => {
  const FILE = "scripts/property-seeds.test.ts";
  const entry = (seed: number) => ({ seed, note: "test", date: "2026-09-27" });
  afterEach(() => {
    overridePinnedSeedsForTesting(undefined);
  });

  test("replay the pinned seeds before the generated runs", () => {
    withEnv({});
    const pinned = [entry(101), entry(202)];
    overridePinnedSeedsForTesting({
      [`${FILE}::replay the pinned seeds before the generated runs`]: pinned,
    });
    const seen: number[] = [];
    const drawnBy = (seed: number): number[] =>
      fc.sample(fc.integer({ min: 0, max: 1_000_000 }), { seed, numRuns: 3 });
    assertProperty(
      fc.property(fc.integer({ min: 0, max: 1_000_000 }), (value) => {
        seen.push(value);
      }),
      { numRuns: 3, seed: 5 },
    );
    expect(seen).toEqual([...drawnBy(101), ...drawnBy(202), ...drawnBy(5)]);
  });

  test("replay the pinned seeds for an async property too", async () => {
    withEnv({});
    overridePinnedSeedsForTesting({
      [`${FILE}::replay the pinned seeds for an async property too`]: [entry(1)],
    });
    let runs = 0;
    await assertProperty(
      fc.asyncProperty(fc.nat(), async () => {
        runs += 1;
      }),
      { numRuns: 2 },
    );
    expect(runs).toBe(2 * 2);
  });

  test("pinned-only assertions never draw a fresh pass or scale its case count", () => {
    withEnv({ PROPERTY_TEST_NUM_RUNS_FACTOR: "3" });
    const pinned = [entry(101), entry(202)];
    overridePinnedSeedsForTesting({
      [`${FILE}::pinned-only assertions never draw a fresh pass or scale its case count`]: pinned,
    });
    let runs = 0;
    assertPinnedProperty(
      fc.property(fc.nat(), () => {
        runs += 1;
      }),
      { numRuns: 200 },
    );
    expect(runs).toBe(pinned.length);
  });

  test("pinned-only assertions fail when no fixed seed is registered", () => {
    withEnv({});
    overridePinnedSeedsForTesting({});
    expect(() => assertPinnedProperty(fc.property(fc.nat(), () => true))).toThrow(
      /No fixed regression seeds/,
    );
  });

  test("explicit examples do not shift a pinned replay path", () => {
    withEnv({});
    overridePinnedSeedsForTesting({
      [`${FILE}::explicit examples do not shift a pinned replay path`]: [
        { ...entry(101), path: "0" },
      ],
    });
    const seen: number[] = [];
    assertProperty(
      fc.property(fc.nat({ max: 10 }), (value) => {
        seen.push(value);
      }),
      { numRuns: 2, seed: 5, examples: [[99]] },
    );
    expect(seen.at(0)).toBe(fc.sample(fc.nat({ max: 10 }), { seed: 101, numRuns: 1 }).at(0));
    expect(seen).toContain(99);
  });

  test("refuse a property with pinned seeds that bypasses assertProperty", () => {
    withEnv({});
    overridePinnedSeedsForTesting({
      [`${FILE}::refuse a property with pinned seeds that bypasses assertProperty`]: [entry(1)],
    });
    expect(() => propertyConfig()).toThrow(/only replay through assertProperty/);
  });

  test(`every ${PROPERTY_SEEDS_FILE} entry names a test that asserts through a registry driver`, () => {
    const problems: string[] = [];
    for (const [key, entries] of Object.entries(readPinnedSeeds())) {
      const [file, title] = key.split("::") as [string, string | undefined];
      let lines: string[] = [];
      try {
        lines = readFileSync(path.join(REPO_ROOT, file), "utf8").split("\n");
      } catch {
        problems.push(`${key}: no such file`);
        continue;
      }
      const sites = lines.flatMap((text, index) =>
        /\bassert(?:Pinned)?Property\(/.test(text) && !/^\s*(?:\*|\/\/|import)/.test(text)
          ? [index + 1]
          : [],
      );
      if (!sites.some((line) => enclosingTitles(lines, line).at(-1) === title)) {
        problems.push(`${key}: no registry driver call inside a test with that title`);
      }
      for (const pinned of entries) {
        if (!Number.isInteger(pinned.seed) || (pinned.seed | 0) !== pinned.seed) {
          problems.push(`${key}: seed ${String(pinned.seed)} is not a 32-bit integer`);
        }
        if (pinned.path !== undefined && !/^\d+(?::\d+)*$/.test(pinned.path)) {
          problems.push(`${key}: path ${JSON.stringify(pinned.path)} is not a fast-check path`);
        }
        if (pinned.note.trim() === "" || !/^\d{4}-\d{2}-\d{2}$/.test(pinned.date)) {
          problems.push(`${key}: every entry needs a note and a YYYY-MM-DD date`);
        }
      }
    }
    expect(problems).toEqual([]);
  });
});

describe("test titles", () => {
  const source = [
    'describe("outer", () => {',
    '  test("sibling", () => {});',
    "",
    "  test(",
    "    `per fixture (${name})`,",
    "    async () => {",
    "      await fc.assert(",
    "        prop,",
    "        propertyConfig({ numRuns: 4 }),",
    "      );",
    "    },",
    "  );",
    "});",
  ];

  test("read the enclosing titles by indentation, as written", () => {
    expect(enclosingTitles(source, 9)).toEqual(["outer", "per fixture (${name})"]);
    expect(enclosingTitles(source, 2)).toEqual(["outer"]);
  });

  test("turn a template title into a -t pattern", () => {
    expect(titlePattern("per fixture (${name})")).toBe("per fixture \\(.*\\)");
    expect(new RegExp(titlePattern("per fixture (${name})")).test("per fixture (a.docx)")).toBe(
      true,
    );
  });
});
