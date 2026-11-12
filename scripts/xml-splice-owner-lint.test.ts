/** Wiring test for the xml-splice-owner lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-xml-splice(no-hand-rolled-splice)";

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

describe("xml splice owner", () => {
  test("rejects every spelling of a region cut out of a part", () => {
    expect(lintFixture("xml-splice-owner.invalid.ts")).toBe(3);
  });

  test("accepts the owner, an append, and slices that meet", () => {
    expect(lintFixture("xml-splice-owner.valid.ts")).toBe(0);
  });
});
