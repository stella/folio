import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import baseline from "../.oxlint-plugins/folio-test-dom.baseline.json";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-test-dom(dom-lease-owner)";
setDefaultTimeout(30_000);

type LintOptions = { files: string[]; allowRules?: string[] };

const lint = ({ files, allowRules = [] }: LintOptions) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.fixtures.config.ts",
      "--no-ignore",
      ...allowRules.flatMap((rule) => ["--allow", rule]),
      ...files,
    ],
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return { exitCode: result.exitCode, output };
};

describe("test DOM lease ownership", () => {
  test("rejects named, namespace, require and dynamic deep provider imports", () => {
    const result = lint({
      files: ["test/__fixtures__/test-dom-deep.invalid.ts"],
    });
    expect(result.exitCode).not.toBe(0);
    expect(result.output.split(RULE_MARKER).length - 1).toBe(4);
  });

  test("does not match lookalike package prefixes or deep type-only imports", () => {
    const result = lint({
      files: ["test/__fixtures__/test-dom-lookalike.valid.ts"],
    });
    expect(result.output).not.toContain(RULE_MARKER);
    expect(result.exitCode, result.output).toBe(0);
  });

  test.each([
    "test-dom-bind.invalid.ts",
    "test-dom-assignment.invalid.ts",
    "test-dom-destructure.invalid.ts",
  ])(
    "rejects provider acquisition before a later invocation in %s",
    (fixture) => {
      const result = lint({ files: [`test/__fixtures__/${fixture}`] });
      expect(result.output).toContain(RULE_MARKER);
    },
  );

  test("does not mistake a shadowing parameter for the imported provider", () => {
    const result = lint({
      files: ["test/__fixtures__/test-dom-shadow.valid.ts"],
      allowRules: ["no-shadow"],
    });
    expect(result.output).not.toContain(RULE_MARKER);
    expect(result.exitCode, result.output).toBe(0);
  });

  test("rejects raw provider references and module acquisitions", () => {
    const result = lint({ files: ["test/__fixtures__/test-dom.invalid.ts"] });
    expect(result.exitCode).not.toBe(0);
    expect(result.output.split(RULE_MARKER).length - 1).toBe(6);
  });

  test("accepts the shared lease and unrelated registry APIs", () => {
    const result = lint({ files: ["test/__fixtures__/test-dom.valid.ts"] });
    expect(result.output).not.toContain(RULE_MARKER);
    expect(result.exitCode, result.output).toBe(0);
  });

  test("the exact legacy counts remain accepted without granting new test files an allowance", () => {
    expect(Object.values(baseline).every((count) => count > 0)).toBe(true);
    const result = lint({ files: Object.keys(baseline) });
    expect(result.exitCode, result.output).toBe(0);
  });

  test.each([
    "test-dom-require-shadow.valid.ts",
    "test-dom-type-only.valid.ts",
  ])(
    "accepts non-provider runtime bindings and type references in %s",
    (fixture) => {
      const result = lint({
        files: [`test/__fixtures__/${fixture}`],
        allowRules: ["no-shadow"],
      });
      expect(result.output).not.toContain(RULE_MARKER);
      expect(result.exitCode, result.output).toBe(0);
    },
  );
});
