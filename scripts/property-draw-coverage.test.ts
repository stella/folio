/**
 * No assertion reads a tally that a property's draw writes.
 *
 * A property's generated pass draws what its seed chooses and may stop after
 * any case, on a failure or on an exploration finding pull request CI reports
 * without failing. A count of the kinds its predicate saw, asserted after the
 * run, then fails on seeds that never drew a kind and on runs that ended
 * early, in whichever pull request happens to run it. Coverage a property must
 * reach is checked on a fixed sample (test/property-coverage.ts), by fixed
 * `examples`, or by a generator that always yields each kind.
 */

import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import { drawCoverageAssertions } from "./lib/property-draw-coverage";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const SCANNED_ROOTS = [
  "packages",
  "scripts",
  "test",
  "tests",
  "parity",
  "benchmarks",
] as const;
const FIXTURES = "test/__fixtures__/";

setDefaultTimeout(scanTestTimeout(120_000));

const readText = (file: string): string => {
  const absolute = path.join(REPO_ROOT, file);
  if (!existsSync(absolute)) panic(`Cannot read ${file}.`);
  // A byte-order mark is not source text.
  return readFileSync(absolute, "utf8").replace(/^﻿/u, "");
};

/** Sources that draw values; the workspace symlinks are not walked. */
const drawingSources = (): string[] =>
  SCANNED_ROOTS.flatMap((root) =>
    existsSync(path.join(REPO_ROOT, root))
      ? [
          ...new Bun.Glob("**/*.{ts,tsx,mts}").scanSync({
            cwd: path.join(REPO_ROOT, root),
            dot: true,
            onlyFiles: true,
          }),
        ].map((file) => `${root}/${file.replaceAll("\\", "/")}`)
      : [],
  )
    .filter(
      (file) =>
        !/(?:^|\/)(?:node_modules|dist)\//u.test(file) &&
        !file.startsWith(FIXTURES),
    )
    .filter((file) =>
      /\bfc\.(?:property|asyncProperty|sample)\b/u.test(readText(file)),
    )
    .toSorted();

/** The lines a fixture marks `// caught`. */
const markedSites = (file: string): string[] =>
  readText(file)
    .split("\n")
    .flatMap((line, index) =>
      line.includes("// caught") ? [`${file}:${String(index + 1)}`] : [],
    );

const sitesIn = (file: string): string[] =>
  drawCoverageAssertions(file, readText(file)).map(({ site }) => site);

describe("assertions over a property's draw", () => {
  test("no source asserts on a tally its property's draw writes", () => {
    const sources = drawingSources();
    expect(sources.length).toBeGreaterThan(100);
    expect(
      sources.flatMap((file) =>
        drawCoverageAssertions(file, readText(file)).map(
          ({ site, tallies }) => `${site} (${tallies.join(", ")})`,
        ),
      ),
    ).toEqual([]);
  });

  test("the deliberate break is caught at every assertion it marks", () => {
    const fixture = `${FIXTURES}property-draw-coverage.invalid.ts`;
    const marked = markedSites(fixture);
    expect(marked.length).toBeGreaterThanOrEqual(13);
    expect(sitesIn(fixture)).toEqual(marked);
  });

  test("the fixed forms pass", () => {
    expect(sitesIn(`${FIXTURES}property-draw-coverage.valid.ts`)).toEqual([]);
  });

  test("names resolve by scope, through helpers and aliases", () => {
    const source = (body: string) =>
      `import fc from "fast-check";\nconst run = (p: unknown) => p;\n${body}`;
    expect(
      drawCoverageAssertions(
        "probe.ts",
        source(
          [
            "const outer = () => { let hits = 0; run(fc.property(fc.nat(), () => { hits++; })); expect(hits).toBe(1); };",
            "const other = () => { const hits = 5; expect(hits).toBe(5); };",
          ].join("\n"),
        ),
      ).map(({ site }) => site),
    ).toEqual(["probe.ts:3"]);
    expect(
      drawCoverageAssertions(
        "probe.ts",
        source(
          [
            "const tally = new Map<string, number>();",
            "const record = (key: string) => tally.set(key, 1);",
            "run(fc.property(fc.string(), (key) => { record(key); }));",
            "const size = tally.size;",
            "expect(size).toBeGreaterThan(0);",
          ].join("\n"),
        ),
      ).map(({ site, tallies }) => `${site} ${tallies.join(",")}`),
    ).toEqual(["probe.ts:7 tally"]);
  });
});
