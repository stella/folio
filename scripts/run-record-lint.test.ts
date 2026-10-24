/** Wiring tests for the source run-record restoration rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-run-record(restore-source-run-record)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string) => {
  const result = Bun.spawnSync(
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
  );
  const output = result.stdout.toString() + result.stderr.toString();
  return {
    exitCode: result.exitCode,
    findings: output.split(RULE_MARKER).length - 1,
  };
};

describe("source run-record restoration", () => {
  test("rejects missing, mismatched, and cross-switch-arm restoration", () => {
    expect(lintFixture("run-record-metadata.invalid.ts")).toEqual({
      exitCode: 1,
      findings: 7,
    });
  });

  test("accepts source restoration and fresh runs without source marks", () => {
    expect(lintFixture("run-record-metadata.valid.ts")).toEqual({
      exitCode: 0,
      findings: 0,
    });
  });
});
