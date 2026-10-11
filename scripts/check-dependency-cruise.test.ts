import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { assertProperty, propertyTestTimeout } from "../test/property-testing";

import {
  dependencyCruiseCoverageIssue,
  dependencyCruiseOutputIssue,
  MINIMUM_CRUISED_MODULES,
} from "./check-dependency-cruise";

setDefaultTimeout(propertyTestTimeout(5_000));

const fixtureDirectories: string[] = [];
afterEach(() => {
  for (const directory of fixtureDirectories.splice(0))
    rmSync(directory, { recursive: true, force: true });
});

describe("dependency-cruiser coverage", () => {
  test("rejects every count below the floor and accepts every count at or above it", () => {
    assertProperty(
      fc.property(fc.integer({ min: 0, max: 100_000 }), (count) => {
        expect(dependencyCruiseCoverageIssue(count) === null).toBe(
          count >= MINIMUM_CRUISED_MODULES,
        );
      }),
      { id: "rejects every count below the floor and accepts every count at or above it" },
    );
  });

  test("a successful tiny cruise is a failure (positive control)", () => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-cruise-floor-"));
    fixtureDirectories.push(directory);
    const entry = path.join(directory, "entry.js");
    writeFileSync(entry, "export const value = 1;\n");
    const result = Bun.spawnSync(
      [
        path.resolve(import.meta.dirname, "../node_modules/.bin/depcruise"),
        "--no-config",
        "--output-type",
        "err",
        entry,
      ],
      { stdout: "pipe", stderr: "pipe" },
    );
    expect(result.exitCode).toBe(0);
    expect(dependencyCruiseOutputIssue(result.stdout.toString())).toContain(
      "cruised 1 modules; expected at least",
    );
  });

  test("accepts the recorded full-graph summary and summaries with violations", () => {
    expect(
      dependencyCruiseOutputIssue(
        "✔ no dependency violations found (2713 modules, 9101 dependencies cruised)\n",
      ),
    ).toBeNull();
    expect(
      dependencyCruiseOutputIssue(
        "x 1 dependency violations (0 errors, 1 warnings). 2713 modules, 9101 dependencies cruised.\n",
      ),
    ).toBeNull();
  });

  test("missing, duplicate, and invalid summaries fail closed", () => {
    for (const output of [
      "",
      "no violations",
      "NaN modules, 0 dependencies cruised",
      "2713 modules, 1 dependencies cruised\n2713 modules, 1 dependencies cruised",
    ]) {
      expect(dependencyCruiseOutputIssue(output)).not.toBeNull();
    }
    for (const count of [undefined, null, "2713", -1, 1.5, NaN, Infinity]) {
      expect(dependencyCruiseCoverageIssue(count)).not.toBeNull();
    }
  });
});
