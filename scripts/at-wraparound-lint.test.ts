/** Wiring test for the at-wraparound lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-at-wraparound(no-computed-negative-index)";

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

describe("at wraparound", () => {
  test("rejects every index computed by subtraction or negation", () => {
    expect(lintFixture("at-wraparound.invalid.ts")).toBe(6);
  });

  test("accepts literal indexes, forward offsets and the helpers", () => {
    expect(lintFixture("at-wraparound.valid.ts")).toBe(0);
  });
});
