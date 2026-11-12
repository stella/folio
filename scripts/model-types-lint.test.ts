/** Wiring test for the model-type completeness lint rules (issue #845). */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const WIDENING_MARKER = "folio-model-types(no-model-intersection-widening)";
const IN_CHECK_MARKER = "folio-model-types(no-in-check-on-model)";
const MARK_MARKER = "folio-model-types(writable-record-mark)";

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
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return {
    widening: output.split(WIDENING_MARKER).length - 1,
    inCheck: output.split(IN_CHECK_MARKER).length - 1,
    unmarked: output.split(MARK_MARKER).length - 1,
  };
};

describe("model type completeness", () => {
  test("rejects a local intersection widening and an in-check read", () => {
    const counts = lintFixture("model-types.invalid.ts");
    // The plain data literal, a brand with data, and a computed string key.
    expect(counts.widening).toBe(3);
    expect(counts.inCheck).toBe(1);
  });

  test("accepts a type predicate, a phantom brand, an unknown-typed in-check, and a direct read", () => {
    const counts = lintFixture("model-types.valid.ts");
    expect(counts.widening).toBe(0);
    expect(counts.inCheck).toBe(0);
  });

  test("rejects a model record declared without the writable mark", () => {
    // A named record, an inline one, an index-signature record, a pick, an
    // array entry, an interface and the two members of a union.
    expect(lintFixture("writable-record-mark.invalid.ts").unmarked).toBe(8);
  });

  test("accepts marked, read-only and branded records and parameter types", () => {
    expect(lintFixture("writable-record-mark.valid.ts").unmarked).toBe(0);
  });
});
