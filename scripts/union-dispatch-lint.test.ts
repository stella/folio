/** Wiring test for the model-union dispatch lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-union-dispatch(exhaustive-model-union-dispatch)";

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

describe("model union dispatch", () => {
  test("rejects a chain of else-if over a union tag", () => {
    expect(lintFixture("union-dispatch.invalid.ts")).toBe(1);
  });

  test("accepts a switch, a chain that ends in a never check, and a two-way test", () => {
    expect(lintFixture("union-dispatch.valid.ts")).toBe(0);
  });
});
