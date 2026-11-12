/** Wiring test for the XML-escaping ownership lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-xml-escaping(no-hand-rolled-xml-escape)";

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

describe("XML escaping ownership", () => {
  test("rejects a writer that escapes XML by hand", () => {
    expect(lintFixture("xml-escaping.invalid.ts")).toBe(3);
  });

  test("accepts a writer that escapes through the owner", () => {
    expect(lintFixture("xml-escaping.valid.ts")).toBe(0);
  });
});
