/** Wiring test for the verbatim-capture ownership lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-verbatim-capture(no-direct-element-to-xml)";

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

describe("verbatim capture ownership", () => {
  test("rejects a parser that serializes a capture itself", () => {
    expect(lintFixture("verbatim-capture.invalid.ts")).toBe(1);
  });

  test("accepts a parser that captures through the owner", () => {
    expect(lintFixture("verbatim-capture.valid.ts")).toBe(0);
  });
});
