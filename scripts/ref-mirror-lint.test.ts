import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-ref-mirrors(no-write-to-render-mirrored-ref)";

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

describe("render-mirrored refs", () => {
  test("rejects imperative writes to a ref reassigned from render scope", () => {
    expect(lintFixture("ref-mirror.invalid.ts")).toEqual({ status: 1, violations: 2 });
  });

  test("accepts state-driven updates and writes to non-mirror refs", () => {
    expect(lintFixture("ref-mirror.valid.ts")).toEqual({ status: 0, violations: 0 });
  });
});
