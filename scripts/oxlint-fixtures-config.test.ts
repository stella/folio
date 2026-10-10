import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import config from "../oxlint.config.ts";
import fixtureConfig from "../oxlint.fixtures.config.ts";

test("fixture lint clears only its fixture ignore and preserves the production rule and warning policy", () => {
  const { ignorePatterns: productionIgnores, ...productionPolicy } = config;
  const { ignorePatterns: fixtureIgnores, ...fixturePolicy } = fixtureConfig;
  expect(productionIgnores).toContain("test/__fixtures__/**");
  expect(fixtureIgnores).toEqual(
    productionIgnores?.filter((pattern) => pattern !== "test/__fixtures__/**"),
  );
  expect(fixturePolicy).toEqual(productionPolicy);
});

test("every oxlint fixture self-test selects the fixture config", () => {
  const root = new URL("../", import.meta.url).pathname;
  const tests = [
    ...new Bun.Glob("scripts/*lint.test.ts").scanSync({ cwd: root }),
    "scripts/oxlint-config-loaders.test.ts",
  ];
  let invocations = 0;
  for (const filename of tests) {
    const source = readFileSync(`${root}/${filename}`, "utf8");
    if (
      !source.includes('"oxlint"') ||
      (!source.includes('"__fixtures__"') && !source.includes('"test/__fixtures__/'))
    )
      continue;
    invocations += 1;
    expect(source, filename).toContain('"oxlint.fixtures.config.ts"');
    expect(source, filename).not.toContain('"oxlint.config.ts"');
  }
  expect(invocations).toBeGreaterThan(0);
});
