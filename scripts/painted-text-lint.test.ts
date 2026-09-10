import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-painted-text(no-direct-text-shape)";

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
  return {
    status: result.exitCode,
    violations: output.split(RULE_MARKER).length - 1,
  };
};

describe("painted text DOM shape", () => {
  test("rejects direct-child assumptions", () => {
    expect(lintFixture("painted-text.invalid.ts")).toEqual({ status: 1, violations: 4 });
  });

  test("accepts descendant text streams", () => {
    expect(lintFixture("painted-text.valid.ts")).toEqual({ status: 0, violations: 0 });
  });
});
