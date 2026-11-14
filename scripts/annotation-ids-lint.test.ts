/** Wiring test for the annotation-id truthiness lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-annotation-ids(no-annotation-id-truthiness)";

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

describe("annotation ids", () => {
  test("rejects every truthiness test on an annotation id", () => {
    expect(lintFixture("annotation-ids.invalid.ts")).toBe(5);
  });

  test("accepts absence tests, id lists and named ids", () => {
    expect(lintFixture("annotation-ids.valid.ts")).toBe(0);
  });
});
