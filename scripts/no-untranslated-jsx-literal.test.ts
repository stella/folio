/**
 * Wiring test for the `no-untranslated-jsx-literal` oxlint plugin.
 *
 * The rule keeps folio-react's user-facing JSX text inside the use-intl
 * catalogs (packages/core/src/i18n/messages): a raw text literal ships
 * English verbatim to all 18 locales. This test runs the real repo config
 * against two committed fixtures in `test/__fixtures__` (covered by the same
 * config override as the packages/react/src TSX sources but excluded from the
 * repo-wide lint via `ignorePatterns`, hence `--no-ignore` here) and asserts
 * the rule fires on a raw literal and stays silent on `t(...)` usage. That
 * guards the whole chain: plugin registration in `jsPlugins`, the override
 * scoping, and the rule's own literal detection.
 */

import { describe, expect, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "no-untranslated-jsx-literal(no-untranslated-jsx-literal)";

const lintFixture = (fixture: string) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.config.ts",
      // The fixtures sit in `ignorePatterns` so `bun run lint` skips their
      // deliberate violation; linting them explicitly needs the ignore off.
      "--no-ignore",
      path.join("test", "__fixtures__", fixture),
    ],
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return {
    output,
    ruleDiagnostics: output.split(RULE_MARKER).length - 1,
  };
};

describe("no-untranslated-jsx-literal", () => {
  test("flags a raw user-facing JSX text literal", () => {
    const { output, ruleDiagnostics } = lintFixture("untranslated.fixture.tsx");
    expect(ruleDiagnostics).toBe(1);
    expect(output).toContain("Save document");
  });

  test("accepts text rendered through useTranslations()/t()", () => {
    const { ruleDiagnostics } = lintFixture("translated.fixture.tsx");
    expect(ruleDiagnostics).toBe(0);
  });
});
