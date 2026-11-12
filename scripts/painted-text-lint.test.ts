import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-painted-text(no-direct-text-shape)";

setDefaultTimeout(subprocessTestTimeout(2));

const lintFixture = (fixture: string, marker = RULE_MARKER) => {
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
  return {
    status: result.exitCode,
    violations: output.split(marker).length - 1,
  };
};

describe("painted text DOM shape", () => {
  test("text geometry consumes positioned advances rather than raw or prefix widths", () => {
    const marker = "folio-painted-text(positioned-advance-owner)";
    expect(lintFixture("positioned-advance.invalid.ts", marker)).toEqual({
      status: 1,
      violations: 6,
    });
    expect(lintFixture("positioned-advance.valid.ts", marker)).toEqual({
      status: 0,
      violations: 0,
    });
  });
  test("rejects direct-child assumptions", () => {
    expect(lintFixture("painted-text.invalid.ts")).toEqual({
      status: 1,
      violations: 4,
    });
  });

  test("accepts descendant text streams", () => {
    expect(lintFixture("painted-text.valid.ts")).toEqual({
      status: 0,
      violations: 0,
    });
  });
});
