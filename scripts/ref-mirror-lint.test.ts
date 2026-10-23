import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";
import { z } from "zod";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_ID = "folio-ref-mirrors(no-write-to-render-mirrored-ref)";

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
      "--format",
      "json",
      path.join("test", "__fixtures__", fixture),
    ],
    { cwd: REPO_ROOT },
  );
  const { diagnostics } = z
    .object({ diagnostics: z.array(z.object({ code: z.string() })) })
    .parse(JSON.parse(result.stdout.toString()));
  return {
    status: result.exitCode,
    ruleIds: diagnostics.map(({ code }) => code),
  };
};

describe("render-mirrored refs", () => {
  test("rejects imperative writes to a ref reassigned from render scope", () => {
    const result = lintFixture("ref-mirror.invalid.ts");
    expect(result.status).toBe(1);
    expect(result.ruleIds.filter((ruleId) => ruleId === RULE_ID)).toEqual([
      RULE_ID,
      RULE_ID,
    ]);
  });

  test("accepts state-driven updates and writes to non-mirror refs", () => {
    expect(lintFixture("ref-mirror.valid.ts")).toEqual({
      status: 0,
      ruleIds: [],
    });
  });
});
