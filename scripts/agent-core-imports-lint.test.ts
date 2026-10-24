import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

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
    { cwd: path.resolve(import.meta.dir, "..") },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  const marker = "folio-layer-boundaries(agents-contract-import)";
  const violations = output.split(marker).length - 1;
  const hasViolations = violations > 0;
  const lintFailed = result.exitCode !== 0;
  if (lintFailed !== hasViolations) {
    panic(
      `oxlint failed without the expected import-boundary diagnostic:\n${output}`,
    );
  }
  return violations;
};

describe("agents core declaration import ownership", () => {
  test("rejects competing facade imports, re-exports and import types", () => {
    expect(lintFixture("agent-core-imports.invalid.ts")).toBe(15);
  });
  test("allows canonical server, defining leaf owners and unrelated prefixes", () => {
    expect(lintFixture("agent-core-imports.valid.ts")).toBe(0);
  });
  test("leaves facade imports outside the agents scope unchanged", () => {
    expect(lintFixture("core-facade-outside-agents.valid.ts")).toBe(0);
  });
});
