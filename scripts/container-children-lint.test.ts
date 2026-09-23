/** Wiring tests for the container child-dispatch lint rules. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const HAND_ROLLED_MARKER = "folio-container-children(no-hand-rolled-child-dispatch)";
const TABLE_MARKER = "folio-container-children(module-level-handler-tables)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string, marker: string) => {
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
  return output.split(marker).length - 1;
};

describe("container child dispatch ownership", () => {
  test("rejects a container walked by a hand-rolled child-name switch", () => {
    expect(lintFixture("container-children.invalid.ts", HAND_ROLLED_MARKER)).toBe(1);
  });

  test("accepts a container walked through the shared dispatcher", () => {
    expect(lintFixture("container-children.valid.ts", HAND_ROLLED_MARKER)).toBe(0);
  });
});

describe("child dispatch tables", () => {
  test("rejects a table written inline or bound to a function local", () => {
    expect(lintFixture("dispatch-handler-tables.invalid.ts", TABLE_MARKER)).toBe(2);
  });

  test("accepts tables bound at module scope, directly or through a member", () => {
    expect(lintFixture("dispatch-handler-tables.valid.ts", TABLE_MARKER)).toBe(0);
  });
});
