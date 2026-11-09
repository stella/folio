/** Wiring test for the painted-positions owner lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-painted-positions(position-attribute-owner)";

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
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return output.split(RULE_MARKER).length - 1;
};

describe("painted positions owner", () => {
  test("rejects every spelling of a painted position outside the owner", () => {
    expect(lintFixture("painted-positions.invalid.ts")).toBe(8);
  });

  test("accepts the owner's readers, writers and selectors", () => {
    expect(lintFixture("painted-positions.valid.ts")).toBe(0);
  });
});
