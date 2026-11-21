/** Wiring test for the math-extremes lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-math-extremes(no-spread-arguments)";

setDefaultTimeout(subprocessTestTimeout());

const lintFixture = (fixture: string) => {
  const result = expectExited(
    runChild(
      [
        "bun",
        "--bun",
        "oxlint",
        "-c",
        "oxlint.fixtures.config.ts",
        "--no-ignore",
        path.join("test", "__fixtures__", fixture),
      ],
      { cwd: REPO_ROOT },
    ),
  );
  const output = `${result.stdout}${result.stderr}`;
  return output.split(RULE_MARKER).length - 1;
};

describe("math extremes", () => {
  test("rejects every way of passing an array as the argument list", () => {
    expect(lintFixture("math-extremes.invalid.ts")).toBe(5);
  });

  test("accepts fixed argument lists and the folding helpers", () => {
    expect(lintFixture("math-extremes.valid.ts")).toBe(0);
  });
});
