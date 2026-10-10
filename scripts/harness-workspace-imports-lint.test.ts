/** Wiring tests for runtime workspace imports in root Playwright and parity harnesses. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-harness-workspaces(no-undeclared-workspace-runtime-import)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string): number => {
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
  if (result.exitCode !== 0 && !output.includes(RULE_MARKER)) {
    throw new Error(`oxlint failed without the expected rule diagnostic:\n${output}`);
  }
  return output.split(RULE_MARKER).length - 1;
};

describe("no-undeclared-workspace-runtime-import", () => {
  test("flags undeclared runtime imports across static, re-export, dynamic and require syntax", () => {
    expect(lintFixture("harness-workspace.invalid.ts")).toBe(6);
  });

  test("allows type-only imports, relative source imports and root-declared workspaces", () => {
    expect(lintFixture("harness-workspace.valid.ts")).toBe(0);
  });
});
