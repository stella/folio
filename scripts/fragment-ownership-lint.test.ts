/** Wiring test for the page-fragment ownership lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-fragment-ownership(no-direct-page-fragment-push)";

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

describe("page fragment ownership", () => {
  test("rejects direct page fragment insertion", () => {
    expect(lintFixture("fragment-ownership.invalid.ts")).toBe(1);
  });

  test("accepts paginator-owned insertion", () => {
    expect(lintFixture("fragment-ownership.valid.ts")).toBe(0);
  });
});
