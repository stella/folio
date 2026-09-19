/** Wiring test for the container child-dispatch lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-container-children(no-hand-rolled-child-dispatch)";

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

describe("container child dispatch ownership", () => {
  test("rejects a container walked by a hand-rolled child-name switch", () => {
    expect(lintFixture("container-children.invalid.ts")).toBe(1);
  });

  test("accepts a container walked through the shared dispatcher", () => {
    expect(lintFixture("container-children.valid.ts")).toBe(0);
  });
});
