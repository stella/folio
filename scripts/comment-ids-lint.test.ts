/** Wiring test for the comment-id truthiness lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-comment-ids(no-comment-id-truthiness)";

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

describe("comment ids", () => {
  test("rejects every truthiness test on a comment id", () => {
    expect(lintFixture("comment-ids.invalid.ts")).toBe(4);
  });

  test("accepts absence tests and comment id lists", () => {
    expect(lintFixture("comment-ids.valid.ts")).toBe(0);
  });
});
