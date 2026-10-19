import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  buildModuleGraph,
  discoverTestSuites,
  selectFastTestSuitesWithReason,
} from "./ci-test-shards";

const temporaryRoots: string[] = [];

setDefaultTimeout(30_000);

const fixture = (files: Record<string, string>) => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-import-closure-"));
  temporaryRoots.push(root);
  for (const [file, contents] of Object.entries(files)) {
    const target = path.join(root, file);
    mkdirSync(path.dirname(target), { recursive: true });
    writeFileSync(target, contents);
  }
  return root;
};

const suite = (files: string[]) => [{ cwd: "packages/app", preloads: [], files }];

afterEach(() => {
  for (const root of temporaryRoots.splice(0)) rmSync(root, { recursive: true, force: true });
});

describe("reverse import closure", () => {
  test("follows relative imports, re-exports, workspace exports, literals, and cycles", () => {
    const root = fixture({
      "packages/lib/package.json": JSON.stringify({
        name: "@stll/lib",
        exports: { "./feature": "./src/feature.ts" },
      }),
      "packages/lib/src/feature.ts": "export const feature = 1;",
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/reexport.ts": 'export { feature } from "@stll/lib/feature";',
      "packages/app/src/lazy.ts": 'export const lazy = import("./reexport");',
      "packages/app/src/cycle-a.ts": 'import "./cycle-b"; export { lazy } from "./lazy";',
      "packages/app/src/cycle-b.ts": 'import "./cycle-a";',
      "packages/app/src/selected.test.ts": 'import "./cycle-a";',
      "packages/app/src/other.test.ts": "export {};",
      "packages/app/src/third.test.ts": "export {};",
    });
    const files = [
      "packages/app/src/selected.test.ts",
      "packages/app/src/other.test.ts",
      "packages/app/src/third.test.ts",
    ];
    const result = selectFastTestSuitesWithReason({
      suites: suite(files),
      changedFiles: ["packages/lib/src/feature.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(result.reason).toBe("closure");
    expect(result.suites.flatMap(({ files: selected }) => selected)).toEqual([files[0]]);
  });

  test("ignores type-only imports without hiding a runtime import", () => {
    const root = fixture({
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/value.ts": "export type T = string; export const value = 1;",
      "packages/app/src/type-only.test.ts": 'import type { T } from "./value"; export {};',
      "packages/app/src/runtime.test.ts":
        'import type { T } from "./value"; import { value } from "./value"; void value;',
      "packages/app/src/other.test.ts": "export {};",
    });
    const result = selectFastTestSuitesWithReason({
      suites: suite([
        "packages/app/src/type-only.test.ts",
        "packages/app/src/runtime.test.ts",
        "packages/app/src/other.test.ts",
      ]),
      changedFiles: ["packages/app/src/value.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(result.suites.flatMap(({ files }) => files)).toEqual([
      "packages/app/src/runtime.test.ts",
    ]);
  });

  test("seeds files with unresolved or non-literal imports for every code change", () => {
    for (const source of ['import "unknown-package";', "import(process.env.MODULE);"]) {
      const root = fixture({
        "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
        "packages/app/src/value.ts": "export {};",
        "packages/app/src/one.test.ts": source,
        "packages/app/src/two.test.ts": "export {};",
        "packages/app/src/three.test.ts": "export {};",
      });
      const result = selectFastTestSuitesWithReason({
        suites: suite([
          "packages/app/src/one.test.ts",
          "packages/app/src/two.test.ts",
          "packages/app/src/three.test.ts",
        ]),
        changedFiles: ["packages/app/src/value.ts"],
        focusedFiles: [],
        repoRoot: root,
      });
      expect(result.reason).toBe("closure");
      expect(result.suites.flatMap(({ files }) => files)).toEqual(["packages/app/src/one.test.ts"]);
      expect(result.cause).toBe("unknown dependencies in 1 file: packages/app/src/one.test.ts");

      const documentationOnly = selectFastTestSuitesWithReason({
        suites: suite([
          "packages/app/src/one.test.ts",
          "packages/app/src/two.test.ts",
          "packages/app/src/three.test.ts",
        ]),
        changedFiles: ["README.md"],
        focusedFiles: [],
        repoRoot: root,
      });
      expect(documentationOnly.suites).toEqual([]);
      expect(documentationOnly.cause).toBeUndefined();
    }
  });

  test("selects every test in a suite when a reached preload depends on the change", () => {
    const root = fixture({
      "packages/app/package.json": JSON.stringify({
        name: "@stll/app",
        scripts: { test: "bun test src" },
      }),
      "packages/app/bunfig.toml": '[test]\npreload = ["./test/setup.ts"]\n',
      "packages/app/src/value.ts": "export {};",
      "packages/app/test/setup.ts": 'import "../src/value";',
      "packages/app/src/one.test.ts": "export {};",
      "packages/app/src/two.test.ts": "export {};",
      "packages/other/package.json": JSON.stringify({
        name: "@stll/other",
        scripts: { test: "bun test src" },
      }),
      "packages/other/src/three.test.ts": "export {};",
      "packages/other/src/four.test.ts": "export {};",
      "packages/other/src/unrelated.ts": "export {};",
    });
    const suites = discoverTestSuites(root);
    expect(suites.find(({ cwd }) => cwd === "packages/app")?.preloads).toEqual(["./test/setup.ts"]);
    const reached = selectFastTestSuitesWithReason({
      suites,
      changedFiles: ["packages/app/src/value.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(reached.suites.flatMap(({ files }) => files)).toEqual([
      "packages/app/src/one.test.ts",
      "packages/app/src/two.test.ts",
    ]);

    const unrelated = selectFastTestSuitesWithReason({
      suites,
      changedFiles: ["packages/other/src/unrelated.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(unrelated.suites).toEqual([]);
  });

  test("fails closed when a changed declaration source is excluded from the graph", () => {
    const file = "packages/app/src/generated.d.ts";
    const root = fixture({
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      [file]: "export declare const generated: string;",
      "packages/app/src/one.test.ts": "export {};",
      "packages/app/src/two.test.ts": "export {};",
    });
    const result = selectFastTestSuitesWithReason({
      suites: suite(["packages/app/src/one.test.ts", "packages/app/src/two.test.ts"]),
      changedFiles: [file],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(result.reason).toBe("full-fallback");
    expect(result.cause).toBe(`changed source is absent from graph: ${file}`);
  });

  test("selects package tests for package data and falls back at 70 percent", () => {
    const root = fixture({
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/value.ts": "export {};",
      "packages/app/src/one.test.ts": 'import "./value";',
      "packages/app/src/two.test.ts": 'import "./value";',
      "packages/app/src/three.test.ts": "export {};",
      "packages/app/fixtures/input.json": "{}",
      "packages/other/package.json": JSON.stringify({ name: "@stll/other" }),
      "packages/other/src/four.test.ts": "export {};",
      "packages/other/src/five.test.ts": "export {};",
    });
    const files = [
      "packages/app/src/one.test.ts",
      "packages/app/src/two.test.ts",
      "packages/app/src/three.test.ts",
    ];
    const data = selectFastTestSuitesWithReason({
      suites: [
        ...suite(files),
        {
          cwd: "packages/other",
          preloads: [],
          files: ["packages/other/src/four.test.ts", "packages/other/src/five.test.ts"],
        },
      ],
      changedFiles: ["packages/app/fixtures/input.json"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(data.reason).toBe("package-data");
    expect(data.suites.flatMap(({ files: selected }) => selected)).toEqual(files);
    const threshold = selectFastTestSuitesWithReason({
      suites: suite(files),
      changedFiles: ["packages/app/src/value.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(threshold.reason).toBe("closure");
    expect(threshold.selectedCount).toBe(2);

    const thresholdFiles = Array.from(
      { length: 10 },
      (_, index) => `packages/app/src/threshold-${index}.test.ts`,
    );
    for (const [index, file] of thresholdFiles.entries()) {
      writeFileSync(path.join(root, file), index < 7 ? 'import "./value";' : "export {};");
    }
    const atThreshold = selectFastTestSuitesWithReason({
      suites: suite(thresholdFiles),
      changedFiles: ["packages/app/src/value.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(atThreshold.reason).toBe("full-fallback");
    expect(atThreshold.cause).toContain("selection reached 7/10 files");
  });
});

test("the reported regression tests are selected from the real repository graph", () => {
  const suites = discoverTestSuites();
  const expected = [
    "packages/core/src/version-comparison.test.ts",
    "packages/core/src/ai-edits/deletedNoteInsertionRejection.test.ts",
    "packages/core/src/prosemirror/plugins/suggestionHistoryRoundTrip.test.ts",
    "packages/core/src/prosemirror/plugins/tableDragInputTrace.test.ts",
    "packages/core/src/__tests__/revisionResolutionReference.test.ts",
    "packages/core/src/compare/compare.property.test.ts",
    "packages/docx-core/src/ops/__tests__/review.property.test.ts",
  ];
  const result = selectFastTestSuitesWithReason({
    suites,
    changedFiles: [
      "packages/core/src/internal/wholeStoryRevisionResolution.ts",
      "packages/docx-core/src/ops/blockDeletion.ts",
      "packages/docx-core/src/ops/documentOps.ts",
      "packages/docx-core/src/ops/resolve.ts",
      "packages/docx-core/src/ops/__tests__/blockDeletion.test.ts",
      "packages/docx-core/src/ops/__tests__/documentOps.test.ts",
      "test/consumer-scenarios/scenarios/paragraph-carriers.test.ts",
    ],
    focusedFiles: [],
  });
  const selected = new Set(result.suites.flatMap(({ files }) => files));
  for (const file of expected) expect(selected.has(file), file).toBe(true);
});

test("random real source subsets preserve changed tests and closure intersections", () => {
  const suites = discoverTestSuites();
  const graph = buildModuleGraph();
  const tests = suites.flatMap(({ files }) => files);
  const sources = [...graph.codeFiles].filter((file) => !/\.test\./u.test(file));
  const dependencies = new Map<string, Set<string>>();
  for (const [dependency, importers] of graph.reverse) {
    for (const importer of importers) {
      const imported = dependencies.get(importer) ?? new Set<string>();
      imported.add(dependency);
      dependencies.set(importer, imported);
    }
  }
  let state = 0x19_62_10_18;
  const random = () => (state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0) / 2 ** 32;
  for (let iteration = 0; iteration < 8; iteration++) {
    const changed = new Set(sources.filter(() => random() < 0.0015));
    changed.add(tests[Math.floor(random() * tests.length)] ?? tests[0]);
    const result = selectFastTestSuitesWithReason({
      suites,
      changedFiles: [...changed],
      focusedFiles: [],
      graph,
    });
    const selected = result.suites.flatMap(({ files }) => files);
    for (const file of changed) if (/\.test\./u.test(file)) expect(selected).toContain(file);
    if (result.reason !== "closure") continue;
    for (const testFile of selected) {
      const seen = new Set([testFile]);
      const pending = [testFile];
      while (pending.length > 0) {
        const importer = pending.pop();
        if (!importer) continue;
        for (const dependency of dependencies.get(importer) ?? []) {
          if (seen.has(dependency)) continue;
          seen.add(dependency);
          pending.push(dependency);
        }
      }
      expect([...changed, ...graph.unresolved.keys()].some((file) => seen.has(file))).toBe(true);
    }
  }
});
