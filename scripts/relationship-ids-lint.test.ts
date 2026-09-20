/** Wiring test for the relationship-id lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-relationship-ids(no-empty-relationship-id)";

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

describe("relationship ids", () => {
  test("rejects every spelling of an empty reference", () => {
    expect(lintFixture("relationship-ids.invalid.ts")).toBe(3);
  });

  test("accepts absence and references that name something", () => {
    expect(lintFixture("relationship-ids.valid.ts")).toBe(0);
  });
});
