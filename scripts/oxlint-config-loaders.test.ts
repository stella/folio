/**
 * The lint config must load under both module loaders that run it.
 *
 * `bun --bun oxlint` (what lefthook runs) resolves an extensionless relative
 * import to a `.ts` file; plain `bunx oxlint` hands the plugins to Node, whose
 * ESM resolver refuses one. A plugin that imports its data extensionless
 * therefore lints locally and dies the moment oxlint runs outside bun, and the
 * failure is a config-parse error, not a lint diagnostic, so every rule goes
 * silent at once. Run the fixtures through both loaders and require the same
 * diagnostics.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const FIXTURES = [
  path.join("test", "__fixtures__", "identity-attributes.invalid.ts"),
  path.join("test", "__fixtures__", "reserved-values.invalid.ts"),
] as const;
const EXPECTED_RULES = [
  "folio-identity-attributes(no-prefix-resolved-identity-read)",
  "folio-reserved-values(no-bare-reserved-compare)",
] as const;

setDefaultTimeout(120_000);

const LOADERS = {
  bun: ["bun", "--bun", "oxlint"],
  node: ["bunx", "oxlint"],
} as const satisfies Record<string, readonly string[]>;

type Loader = keyof typeof LOADERS;

type LintResult = { exitCode: number; output: string };

const lintFixtures = (loader: Loader): LintResult => {
  const result = Bun.spawnSync(
    [...LOADERS[loader], "-c", "oxlint.config.ts", "--no-ignore", ...FIXTURES],
    {
      cwd: REPO_ROOT,
    },
  );
  return {
    exitCode: result.exitCode,
    output: `${result.stdout.toString()}${result.stderr.toString()}`,
  };
};

/** `rule(name)` markers, sorted: the diagnostics both loaders must agree on. */
const rulesReported = (output: string): string[] =>
  [...output.matchAll(/\s(error|warning)\s+([\w-]+\([\w-]+\))/g)].map(([, , rule]) => rule).sort();

const outputs = new Map<Loader, LintResult>(
  (Object.keys(LOADERS) as Loader[]).map((loader) => [loader, lintFixtures(loader)]),
);

describe("oxlint config", () => {
  for (const loader of Object.keys(LOADERS) as Loader[]) {
    test(`loads every JS plugin under the ${loader} loader`, () => {
      const { exitCode, output } = outputs.get(loader) ?? { exitCode: -1, output: "" };
      expect(exitCode).toBe(1);
      expect(output).not.toContain("Failed to load JS plugin");
      expect(output).not.toContain("Failed to parse oxlint configuration file");
    });
  }

  test("reports the same rules through both loaders", () => {
    const node = rulesReported(outputs.get("node")?.output ?? "");
    expect(new Set(node)).toEqual(new Set(EXPECTED_RULES));
    expect(node).toEqual(rulesReported(outputs.get("bun")?.output ?? ""));
  });
});
