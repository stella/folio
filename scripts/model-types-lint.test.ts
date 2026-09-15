/** Wiring test for the model-type completeness lint rules (issue #845). */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const WIDENING_MARKER = "folio-model-types(no-model-intersection-widening)";
const IN_CHECK_MARKER = "folio-model-types(no-in-check-on-model)";

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
    widening: output.split(WIDENING_MARKER).length - 1,
    inCheck: output.split(IN_CHECK_MARKER).length - 1,
  };
};

describe("model type completeness", () => {
  test("rejects a local intersection widening and an in-check read", () => {
    const counts = lintFixture("model-types.invalid.ts");
    expect(counts.widening).toBe(1);
    expect(counts.inCheck).toBe(1);
  });

  test("accepts a type predicate, an unknown-typed in-check, and a direct read", () => {
    const counts = lintFixture("model-types.valid.ts");
    expect(counts.widening).toBe(0);
    expect(counts.inCheck).toBe(0);
  });
});
