/** Wiring test for the relationship-id lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-relationship-ids(no-empty-relationship-id)";

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

describe("relationship ids", () => {
  test("rejects every spelling of an empty reference", () => {
    expect(lintFixture("relationship-ids.invalid.ts")).toBe(3);
  });

  test("accepts absence and references that name something", () => {
    expect(lintFixture("relationship-ids.valid.ts")).toBe(0);
  });
});
