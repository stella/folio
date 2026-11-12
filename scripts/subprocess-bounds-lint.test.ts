/** Wiring test for the lint rule that bounds child-process tests by their children. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-subprocess-bounds(bounded-child-tests)";

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
        path.join("test", "__fixtures__", "subprocess-bounds", fixture),
      ],
      { cwd: REPO_ROOT },
    ),
  );
  const output = `${result.stdout}${result.stderr}`;
  expect(output).not.toContain("Failed to load");
  expect(output).not.toContain("Failed to parse");
  return output.split(RULE_MARKER).length - 1;
};

describe("subprocess test bounds", () => {
  test("rejects bare test, default, child and hang limits and a smaller budget", () => {
    expect(lintFixture("bare-limits.invalid.ts")).toBe(6);
  });

  test("rejects tests and hooks that start a child under the runner's default", () => {
    expect(lintFixture("runner-default.invalid.ts")).toBe(4);
  });

  test("rejects a later default that replaces the subprocess budget", () => {
    expect(lintFixture("later-default.invalid.ts")).toBe(2);
  });

  test("rejects a scan budget as the default for a child", () => {
    expect(lintFixture("scan-default.invalid.ts")).toBe(1);
  });

  test("accepts the subprocess budget as default, per test and in options", () => {
    expect(lintFixture("bounded.valid.ts")).toBe(0);
  });

  test("accepts a per-test budget under another default", () => {
    expect(lintFixture("per-test.valid.ts")).toBe(0);
  });

  test("leaves files without a child alone", () => {
    expect(lintFixture("no-child.valid.ts")).toBe(0);
  });
});
