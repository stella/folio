/** Wiring test for the compile-time-proof placement lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-typecheck-proofs(no-type-suppression-in-test)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.config.ts",
      "--no-ignore",
      path.join("test", "__fixtures__", fixture),
    ],
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return output.split(RULE_MARKER).length - 1;
};

describe("compile-time proof placement", () => {
  test("rejects every suppression directive, in line and block comments", () => {
    expect(lintFixture("typecheck-proofs.invalid.ts")).toBe(4);
  });

  test("leaves prose that names a directive alone", () => {
    expect(lintFixture("typecheck-proofs.valid.ts")).toBe(0);
  });
});
