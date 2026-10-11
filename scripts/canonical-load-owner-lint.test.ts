import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";
import { z } from "zod";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_ID = "folio-canonical-load(evaluation-owner)";
const FIXTURE_CASES = {
  "canonical-load.direct.invalid.ts": [RULE_ID, RULE_ID, RULE_ID],
  "canonical-load.aliases.invalid.ts": [
    RULE_ID,
    RULE_ID,
    RULE_ID,
    RULE_ID,
    RULE_ID,
    RULE_ID,
    RULE_ID,
  ],
  "canonical-load.valid.ts": [],
  "canonical-load.scopes.valid.ts": [],
  "canonical-load.scopes.invalid.ts": [RULE_ID, RULE_ID, RULE_ID, RULE_ID],
};

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
  return { status: result.exitCode, ruleIds: diagnostics.map(({ code }) => code) };
};

describe("canonical load evaluation ownership", () => {
  test("exercises every declared canonical load fixture", () => {
    const fixtures = Array.from(
      new Bun.Glob("canonical-load*.ts").scanSync({
        cwd: path.join(REPO_ROOT, "test", "__fixtures__"),
      }),
    );
    expect(Object.keys(FIXTURE_CASES).sort()).toEqual(fixtures.sort());
  });

  for (const [fixture, ruleIds] of Object.entries(FIXTURE_CASES))
    test(fixture, () => {
      expect(lintFixture(fixture)).toEqual({ status: ruleIds.length === 0 ? 0 : 1, ruleIds });
    });
});
