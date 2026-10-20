import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { propertyConfig, propertyTestTimeout } from "../test/property-testing";
import {
  aggregateTestTimings,
  createGraphqlRequest,
  selectTimingArtifacts,
} from "./aggregate-ci-test-timings";
import {
  assignTestShards,
  buildModuleGraph,
  discoverTestSuites,
  selectFastTestSuitesWithReason,
  shardLoads,
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
  test("scans modules reached outside the graph roots", () => {
    const root = fixture({
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/selected.test.ts": 'import "../../../specifications/registry";',
      "packages/app/src/other.test.ts": "export {};",
      "packages/app/src/third.test.ts": "export {};",
      "specifications/registry.ts": 'import "./colors";',
      "specifications/colors.ts": "export const colors = [];",
    });
    const result = selectFastTestSuitesWithReason({
      suites: suite([
        "packages/app/src/selected.test.ts",
        "packages/app/src/other.test.ts",
        "packages/app/src/third.test.ts",
      ]),
      changedFiles: ["specifications/colors.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(result.reason).toBe("closure");
    expect(result.suites.flatMap(({ files }) => files)).toEqual([
      "packages/app/src/selected.test.ts",
    ]);
  });

  test("fails closed for an unreachable code file outside the graph roots", () => {
    const file = "specifications/unreachable.ts";
    const root = fixture({
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/one.test.ts": "export {};",
      "packages/app/src/two.test.ts": "export {};",
      [file]: "export const unreachable = true;",
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
    for (const source of [
      'import "unknown-package";',
      "import(process.env.MODULE);",
      "const name = 'value'; import(`./${name}`);",
      'const name = "value"; import("./handlers/" + name);',
      "const value = `${import(process.env.MODULE)}`;",
    ]) {
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

  test("resolves a plain template literal dynamic import", () => {
    const root = fixture({
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/value.ts": "export {};",
      "packages/app/src/one.test.ts": "import(`./value`);",
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
    expect(result.suites.flatMap(({ files }) => files)).toEqual(["packages/app/src/one.test.ts"]);
    expect(result.cause).toBeUndefined();
  });

  test("follows a tsconfig path alias that resolves inside the repository", () => {
    const root = fixture({
      "tsconfig.json": JSON.stringify({
        compilerOptions: { baseUrl: ".", paths: { "@app/*": ["packages/app/src/*"] } },
      }),
      "packages/app/package.json": JSON.stringify({ name: "@stll/app" }),
      "packages/app/src/value.ts": "export const value = 1;",
      "packages/app/src/one.test.ts": 'import { value } from "@app/value"; void value;',
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
    expect(result.cause).toBeUndefined();
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

  test("scans a suite preload outside the graph roots and follows its imports", () => {
    const root = fixture({
      "packages/app/package.json": JSON.stringify({
        name: "@stll/app",
        scripts: { test: "bun test --preload ./setup.ts src" },
      }),
      "packages/app/setup.ts": 'import "./src/value";',
      "packages/app/src/value.ts": "export {};",
      "packages/app/src/one.test.ts": "export {};",
      "packages/app/src/two.test.ts": "export {};",
    });
    const suites = discoverTestSuites(root);
    const result = selectFastTestSuitesWithReason({
      suites,
      changedFiles: ["packages/app/src/value.ts"],
      focusedFiles: [],
      repoRoot: root,
    });
    expect(result.suites.flatMap(({ files }) => files)).toEqual([
      "packages/app/src/one.test.ts",
      "packages/app/src/two.test.ts",
    ]);
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

const temporaryDirectories: string[] = [];

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) rmSync(directory, { recursive: true });
});

describe("measured test shard balancing", () => {
  test(
    "longest-first packing stays within 1.25x mean when no file exceeds the mean",
    () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 1_000 }), { minLength: 100, maxLength: 100 }),
          (weights) => {
            const timings = Object.fromEntries(
              weights.map((weight, index) => [`test-${index}`, weight]),
            );
            const mean = weights.reduce((sum, weight) => sum + weight, 0) / 4;
            fc.pre(Math.max(...weights) <= mean);
            const loads = shardLoads(assignTestShards(Object.keys(timings), timings), timings);
            expect(Math.max(...loads)).toBeLessThanOrEqual(mean * 1.25);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(5_000),
  );

  test("unmeasured files use the median measured weight", () => {
    const timings = { a: 10, b: 30, c: 50 };
    expect(shardLoads(assignTestShards(["a", "b", "c", "new"], timings), timings)).toEqual([
      50, 30, 30, 10,
    ]);
  });

  test("aggregation takes per-run medians, follows discovery, and is byte-stable", () => {
    const artifacts = [
      { runId: "one", timings: { kept: 10, vanished: 1 } },
      { runId: "two", timings: { kept: 30 } },
      { runId: "three", timings: { kept: 20 } },
    ];
    const first = aggregateTestTimings(artifacts, ["new", "kept"]);
    expect(first).toBe(aggregateTestTimings(artifacts.toReversed(), ["kept", "new"]));
    expect(JSON.parse(first)).toEqual({ kept: 20, new: 15 });
    expect(first).not.toContain("vanished");
  });

  test("timing artifact discovery distinguishes a failed listing from no data", async () => {
    const failure = new Error("run listing failed");
    expect(selectTimingArtifacts("stella/folio", async () => Promise.reject(failure))).rejects.toBe(
      failure,
    );
    expect(
      await selectTimingArtifacts("stella/folio", async () => ({ workflow_runs: [] })),
    ).toEqual([]);
  });

  test("timing artifact discovery selects complete sets from eligible runs", async () => {
    const fetchJson = async (url: string) => {
      if (url.includes("/workflows/")) {
        return {
          workflow_runs: [
            { event: "pull_request", id: 1 },
            { event: "push", id: 2 },
          ],
        };
      }
      return {
        artifacts: [1, 2, 3, 4].map((id) => ({ id: 20 + id, name: `test-timings-2-${id}` })),
      };
    };
    expect(await selectTimingArtifacts("stella/folio", fetchJson)).toEqual(
      [1, 2, 3, 4].map((id) => ({ artifactId: 20 + id, runId: 2 })),
    );
  });

  test("timing artifact discovery reads later pages when the first holds no eligible run", async () => {
    const fetchJson = async (url: string) => {
      if (url.includes("/workflows/")) {
        if (url.endsWith("&page=1")) {
          return {
            workflow_runs: Array.from({ length: 100 }, (_, index) => ({
              event: "pull_request",
              id: 1_000 + index,
            })),
          };
        }
        if (url.endsWith("&page=2")) return { workflow_runs: [{ event: "push", id: 5 }] };
        return { workflow_runs: [] };
      }
      return {
        artifacts: [1, 2, 3, 4].map((id) => ({ id: 50 + id, name: `test-timings-5-${id}` })),
      };
    };
    expect(await selectTimingArtifacts("stella/folio", fetchJson)).toEqual(
      [1, 2, 3, 4].map((id) => ({ artifactId: 50 + id, runId: 5 })),
    );
  });

  test("timing artifact discovery skips runs whose artifacts have expired", async () => {
    const fetchJson = async (url: string) => {
      if (url.includes("/workflows/")) {
        return {
          workflow_runs: [
            { event: "push", id: 3 },
            { event: "merge_group", id: 4 },
          ],
        };
      }
      const runId = url.includes("/runs/3/") ? 3 : 4;
      return {
        artifacts: [1, 2, 3, 4].map((id) => ({
          id: runId * 10 + id,
          name: `test-timings-${runId}-${id}`,
          expired: runId === 3,
        })),
      };
    };
    expect(await selectTimingArtifacts("stella/folio", fetchJson)).toEqual(
      [1, 2, 3, 4].map((id) => ({ artifactId: 40 + id, runId: 4 })),
    );
  });

  test("assignment is deterministic for identical inputs", () => {
    const timings = { a: 10, b: 10, c: 5, d: 5, e: 1 };
    const assigned = () =>
      assignTestShards(Object.keys(timings), timings).map((shard) => Array.from(shard));
    expect(assigned()).toEqual(assigned());
  });

  test("writes large timing contents into a GraphQL request without argv transport", () => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-timing-request-"));
    temporaryDirectories.push(directory);
    const timingFile = path.join(directory, "timings.json");
    const requestFile = path.join(directory, "request.json");
    const timingBytes = Buffer.from(`${JSON.stringify({ sample: "x".repeat(210_000) })}\n`);
    writeFileSync(timingFile, timingBytes);

    const args = [
      "bun",
      path.join(import.meta.dir, "aggregate-ci-test-timings.ts"),
      "--graphql-request",
      requestFile,
      "--timings",
      timingFile,
      "--repository",
      "stella/folio",
      "--branch",
      "ci/test-timings",
      "--base",
      "0123456789abcdef0123456789abcdef01234567",
    ];
    expect(Math.max(...args.map((argument) => Buffer.byteLength(argument)))).toBeLessThan(4_096);
    const result = Bun.spawnSync(args, { stderr: "pipe" });
    expect(result.exitCode, result.stderr.toString()).toBe(0);

    const request = JSON.parse(readFileSync(requestFile, "utf8"));
    expect(request).toEqual(
      createGraphqlRequest({
        base: "0123456789abcdef0123456789abcdef01234567",
        branch: "ci/test-timings",
        repository: "stella/folio",
        timingFile,
      }),
    );
    expect(
      Buffer.from(request.variables.input.fileChanges.additions[0].contents, "base64"),
    ).toEqual(timingBytes);

    const workflow = readFileSync(
      path.resolve(import.meta.dir, "../.github/workflows/test-timing-refresh.yml"),
      "utf8",
    );
    expect(workflow).toContain('--graphql-request "$request"');
    expect(workflow).not.toContain("--arg contents");
    expect(workflow).not.toContain("< <(");
  });
});
