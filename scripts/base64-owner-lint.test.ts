/** Wiring test for the base64-owner lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-base64(no-hand-rolled-base64)";

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

describe("base64 owner", () => {
  test("rejects every spelling of a hand-rolled encoder", () => {
    expect(lintFixture("base64-owner.invalid.ts")).toBe(3);
  });

  test("accepts the owner and leaves decoding alone", () => {
    expect(lintFixture("base64-owner.valid.ts")).toBe(0);
  });
});
