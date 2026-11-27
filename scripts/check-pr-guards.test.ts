import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { generatedChecks, selectPrGuards } from "./check-pr-guards";

const ROOT = path.resolve(import.meta.dir, "..");
const scripts = JSON.parse(
  readFileSync(path.join(ROOT, "package.json"), "utf8"),
).scripts as Record<string, string>;

const ids = (files: readonly string[]): string[] =>
  selectPrGuards(files, scripts).map((guard) => guard.id);

describe("check:pr-guards selection", () => {
  test("a documentation change runs no guard", () => {
    expect(ids(["README.md"])).toEqual([]);
  });

  test("a package source change runs the source guards over its own files", () => {
    const files = [
      "packages/docx-core/src/ops/apply.ts",
      "packages/core/src/canonical/canonicalSession.ts",
      "README.md",
    ];
    const guards = selectPrGuards(files, scripts);
    const selected = guards.map((guard) => guard.id);
    for (const id of [
      "check:normalization-warnings",
      "check:canonical-cutover",
      "check:strict-name-tolerance",
      "check:reserved-value-coverage",
      "check:reserved-values",
      "check:identity-attributes",
      ...generatedChecks(scripts).filter(
        (name) => name !== "generate:joining-types:check",
      ),
    ])
      expect(selected).toContain(id);
    expect(selected).not.toContain("check:i18n");
    expect(selected).not.toContain("check:dedupe");
    const reserved = guards.find(
      (guard) => guard.id === "check:reserved-values",
    );
    expect(reserved?.command).toEqual([
      "bun",
      "scripts/reserved-value-baseline.ts",
      "--check",
      "--files",
      "packages/docx-core/src/ops/apply.ts",
      "packages/core/src/canonical/canonicalSession.ts",
    ]);
  });

  test("the Unicode tables are checked only when their own inputs change", () => {
    expect(ids(["packages/core/src/docx/paragraphParser.ts"])).not.toContain(
      "generate:joining-types:check",
    );
    expect(ids(["packages/core/src/utils/joiningTypes.gen.ts"])).toContain(
      "generate:joining-types:check",
    );
  });

  test("a manifest change runs the lockfile checks", () => {
    for (const file of ["bun.lock", "packages/docx-core/package.json"])
      expect(ids([file])).toEqual([
        "check:lockfile-versions",
        "check:lockfile-overrides",
        "check:lockfile-trust",
        "check:dedupe",
      ]);
  });

  test("an adapter component runs the parity and catalog checks", () => {
    const selected = ids(["packages/react/src/components/Toolbar.tsx"]);
    for (const id of [
      "check:i18n",
      "check:parity-contract",
      "check:export-parity",
    ])
      expect(selected).toContain(id);
  });

  test("every guard run through a package script names one that exists", () => {
    const everything = [
      "bun.lock",
      "packages/react/src/a.tsx",
      "packages/core/src/utils/joiningTypes.gen.ts",
      "test/corpus/lock.json",
    ];
    for (const guard of selectPrGuards(everything, scripts))
      if (guard.command[1] === "run")
        expect(Object.hasOwn(scripts, guard.id)).toBe(true);
  });
});
