import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { seedFileFor } from "../test/seed-registry";
import { subprocessBudgetBaseMs } from "../test/subprocess";
import { parseSourceModule, walkAst } from "./lib/oxc-ast";
import {
  areaOf,
  EXTENDED_PROPERTY_SHARD_BUDGET_MS,
  planExtendedPropertyShards,
  propertyBatches,
  propertyFileProfile,
  propertyFiles,
  propertyShardTimings,
  ROOT_TEST_DIRS,
  runPropertyBatches,
  selectPropertyFiles,
  shardPropertyFiles,
  touchedAreas,
} from "./property-areas";

const makePropertyFile = (repoPath: string, weightMs = 5_000) => ({
  area: areaOf(repoPath) as string,
  packageDir: repoPath.split("/").slice(0, 2).join("/"),
  file: repoPath,
  weightMs,
});

const ALL = [
  makePropertyFile("packages/core/src/ai-edits/batchOverlap.test.ts"),
  makePropertyFile("packages/core/src/compare/compare.property.test.ts"),
  makePropertyFile("packages/core/src/docx/xmlSerialize.property.test.ts"),
  makePropertyFile("packages/docx-core/src/markdown/markdown.test.ts"),
];

/** Independently inventory actual calls, excluding matches in comments and strings. */
const actualPropertyDrivers = (
  roots: readonly string[] = ["packages", ...ROOT_TEST_DIRS],
): string[] => {
  const root = path.resolve(import.meta.dir, "..");
  const files = roots.flatMap((dir) =>
    [
      ...new Bun.Glob("**/*.{ts,tsx}").scanSync({
        cwd: path.join(root, dir),
        absolute: true,
      }),
    ].filter(
      (file) =>
        !path.relative(root, file).split(path.sep).includes("node_modules"),
    ),
  );
  return files
    .filter((file) => /\.test\.tsx?$/.test(file))
    .filter((file) => {
      const { program } = parseSourceModule(file, readFileSync(file, "utf8"));
      let drives = false;
      walkAst(program, (node) => {
        if (drives) return false;
        if (node.type === "CallExpression") {
          const { callee } = node;
          drives =
            (callee.type === "Identifier" &&
              callee.name === "assertProperty") ||
            (callee.type === "MemberExpression" &&
              !callee.computed &&
              callee.object.type === "Identifier" &&
              callee.object.name === "fc" &&
              callee.property.type === "Identifier" &&
              (callee.property.name === "assert" ||
                callee.property.name === "check"));
        }
        return !drives;
      });
      return drives;
    })
    .map((file) => path.relative(root, file).replaceAll("\\", "/"))
    .toSorted((a, b) => a.localeCompare(b));
};

describe("property areas", () => {
  test("every root with a property driver is declared", () => {
    const root = path.resolve(import.meta.dir, "..");
    const directories = readdirSync(root, { withFileTypes: true })
      .filter(
        (entry) =>
          entry.isDirectory() &&
          !entry.name.startsWith(".") &&
          entry.name !== "node_modules" &&
          entry.name !== "packages",
      )
      .map(({ name }) => name);
    const rootsWithDrivers = directories.filter(
      (directory) => actualPropertyDrivers([directory]).length > 0,
    );
    expect(rootsWithDrivers.toSorted()).toEqual([...ROOT_TEST_DIRS].toSorted());
  });

  test("the workflow matrix exercises every selected file and collects every failure", () => {
    const directory = path.resolve(import.meta.dir, "../.github/workflows");
    const workflows = readdirSync(directory)
      .filter((file) => /\.ya?ml$/u.test(file))
      .map((file) =>
        Bun.YAML.parse(readFileSync(path.join(directory, file), "utf8")),
      )
      .filter((workflow) => workflow.jobs?.["property-areas"]);
    expect(workflows).toHaveLength(1);
    const workflow = workflows.at(0);
    expect(Object.keys(workflow.on)).toEqual(["workflow_dispatch"]);
    const job = workflow.jobs["property-areas"];
    const files = propertyFiles();
    const exercised = [...job.strategy.matrix.shard].flatMap((index: number) =>
      shardPropertyFiles(files, {
        index,
        total: job.strategy.matrix.shard.length,
      }),
    );
    expect(exercised.length).toBe(files.length);
    expect(new Set(exercised).size).toBe(files.length);
    expect(exercised.toSorted((a, b) => a.file.localeCompare(b.file))).toEqual(
      files,
    );
    expect(job.strategy["fail-fast"]).toBe(false);
    const step = job.steps.find(
      (entry: { name?: string }) =>
        entry.name === "Changed areas' properties at 5x numRuns",
    );
    expect(step.run).toContain(
      '--factor 5 --shard "${PROPERTY_SHARD}/${PROPERTY_SHARD_COUNT}"',
    );
    expect(step.env.PROPERTY_SHARD).toBe("${{ matrix.shard }}");
  });

  test("shards cover every selected file exactly once regardless of input order", () => {
    const inventory = propertyFiles();
    for (const total of [1, 2, 4, inventory.length + 1]) {
      const exercised: string[] = [];
      for (let index = 1; index <= total; index += 1) {
        const shard = { index, total };
        const files = shardPropertyFiles(inventory, shard);
        expect(shardPropertyFiles(inventory.toReversed(), shard)).toEqual(
          files,
        );
        // The planner's weights: measured where recorded, else budgets.
        const { timings } = propertyShardTimings(inventory, 1);
        const weights = Object.values(timings);
        const totalWeight = weights.reduce((sum, weight) => sum + weight, 0);
        const heaviest = Math.max(...weights);
        expect(
          files.reduce((sum, { file }) => sum + (timings[file] ?? 0), 0),
        ).toBeLessThanOrEqual(Math.ceil(totalWeight / total) + heaviest);
        for (const { file } of files) exercised.push(file);
      }
      expect(exercised.toSorted()).toEqual(
        inventory.map(({ file }) => file).toSorted(),
      );
      expect(new Set(exercised).size).toBe(exercised.length);
    }
    expect(shardPropertyFiles([], { index: 1, total: 4 })).toEqual([]);
  });

  test("weighted shards spread heavy files with stable filename and shard ties", () => {
    const files = [
      makePropertyFile(
        "packages/docx-core/src/ops/apply.property.test.ts",
        2_400_000,
      ),
      makePropertyFile(
        "packages/docx-core/src/ops/trackedTables.property.test.ts",
        2_400_000,
      ),
      ...Array.from({ length: 20 }, (_, index) =>
        makePropertyFile(
          `packages/core/src/docx/light-${String(index).padStart(2, "0")}.test.ts`,
          30_000,
        ),
      ),
    ];
    const shards = Array.from({ length: 4 }, (_, position) =>
      shardPropertyFiles(files, { index: position + 1, total: 4 }),
    );
    const total = files.reduce((sum, file) => sum + file.weightMs, 0);
    const heaviest = Math.max(...files.map(({ weightMs }) => weightMs));
    for (const [position, shard] of shards.entries()) {
      expect(
        shardPropertyFiles(files.toReversed(), {
          index: position + 1,
          total: 4,
        }),
      ).toEqual(shard);
      expect(
        shard.reduce((sum, file) => sum + file.weightMs, 0),
      ).toBeLessThanOrEqual(Math.ceil(total / 4) + heaviest);
    }
    expect(
      shards.at(0)?.some(({ file }) => file.endsWith("apply.property.test.ts")),
    ).toBe(true);
    expect(
      shards
        .at(1)
        ?.some(({ file }) => file.endsWith("trackedTables.property.test.ts")),
    ).toBe(true);
    expect(
      shards
        .flatMap((shard) => shard)
        .map(({ file }) => file)
        .toSorted(),
    ).toEqual(files.map(({ file }) => file).toSorted());
    const equal = ALL.map((file) =>
      Object.assign({}, file, { weightMs: 5_000 }),
    );
    for (const [position, file] of equal
      .toSorted((a, b) => a.file.localeCompare(b.file))
      .entries())
      // Unmeasured, so the stated budgets tie and filenames break the tie.
      expect(
        shardPropertyFiles(equal, { index: position + 1, total: 4 }, 1, {}),
      ).toEqual([file]);
  });

  test("the expensive operation properties are separated when their area is selected", () => {
    const inventory = propertyFiles();
    const selected = selectPropertyFiles(
      ["packages/docx-core/src/ops/apply.ts"],
      inventory,
    );
    const shards = Array.from({ length: 4 }, (_, position) =>
      shardPropertyFiles(selected, { index: position + 1, total: 4 }),
    );
    const bucketOf = (file: string) =>
      shards.findIndex((shard) => shard.some((entry) => entry.file === file));
    const applyBucket = bucketOf(
      "packages/docx-core/src/ops/__tests__/apply.property.test.ts",
    );
    const trackedBucket = bucketOf(
      "packages/docx-core/src/ops/__tests__/trackedTables.property.test.ts",
    );
    expect(applyBucket).toBeGreaterThanOrEqual(0);
    expect(trackedBucket).toBeGreaterThanOrEqual(0);
    expect(applyBucket).not.toBe(trackedBucket);
  });

  test("invalid scheduling weights fail rather than silently skewing shards", () => {
    for (const weight of [0, -1, Number.NaN, Number.POSITIVE_INFINITY])
      expect(() =>
        shardPropertyFiles(
          [makePropertyFile("packages/core/src/docx/a.test.ts", weight)],
          {
            index: 1,
            total: 4,
          },
        ),
      ).toThrow();
  });

  test("invalid shards fail rather than silently skipping properties", () => {
    for (const shard of [
      { index: 0, total: 4 },
      { index: 5, total: 4 },
      { index: 1, total: 0 },
      { index: 1.5, total: 4 },
      { index: 1, total: Number.NaN },
    ])
      expect(() => shardPropertyFiles(ALL, shard)).toThrow();
  });

  test("an area is the package and the first directory under src", () => {
    expect(areaOf("packages/core/src/compare/diff/align.ts")).toBe(
      "core/compare",
    );
    expect(areaOf("packages/docx-core/src/index.ts")).toBe("docx-core");
    expect(areaOf("scripts/property-areas.ts")).toBeUndefined();
  });

  test("changed source selects its area and the areas coupled to it", () => {
    expect(
      [
        ...touchedAreas(["packages/core/src/internal/compare/blocks.ts"]),
      ].toSorted(),
    ).toEqual([
      "core/ai-edits",
      "core/compare",
      "core/internal",
      "core/prosemirror",
    ]);
    expect(
      selectPropertyFiles(
        ["packages/docx-core/src/markdown/compile.ts"],
        ALL,
      ).map((f) => f.file),
    ).toEqual(["packages/docx-core/src/markdown/markdown.test.ts"]);
  });

  test("a changed test file reruns itself, not its area", () => {
    expect(
      selectPropertyFiles(
        [
          "packages/core/src/compare/compare.property.test.ts",
          "packages/core/src/docx/a.test.ts",
        ],
        ALL,
      ).map((f) => f.file),
    ).toEqual(["packages/core/src/compare/compare.property.test.ts"]);
  });

  test("a change to the pinned seeds reruns the files they name", () => {
    expect(
      selectPropertyFiles(
        [seedFileFor("packages/core/src/ai-edits/batchOverlap.test.ts")],
        ALL,
      ).map((f) => f.file),
    ).toEqual(["packages/core/src/ai-edits/batchOverlap.test.ts"]);
    expect(selectPropertyFiles(["README.md"], ALL)).toEqual([]);
  });

  test("finds the property files of every package", () => {
    const found = propertyFiles();
    expect(
      found.some(({ file: repoPath }) =>
        repoPath.startsWith("packages/core/src/"),
      ),
    ).toBe(true);
    expect(
      found.some(({ file: repoPath }) =>
        repoPath.startsWith("packages/docx-core/src/"),
      ),
    ).toBe(true);
    expect(found.every(({ area }) => area !== undefined)).toBe(true);
    expect(
      found.some(
        ({ file: repoPath }) =>
          repoPath === "scripts/container-survival.property.test.ts",
      ),
    ).toBe(true);
    expect(found.map(({ file }) => file)).toEqual(actualPropertyDrivers());
  });

  test("a root script change selects its properties", () => {
    const rootProperty = {
      area: "scripts",
      packageDir: ".",
      file: "scripts/container-survival.property.test.ts",
      weightMs: 5_000,
    };
    expect(
      selectPropertyFiles(
        ["scripts/container-survival-census.ts"],
        [rootProperty],
      ),
    ).toEqual([rootProperty]);
  });

  test("source changes select every property driver in their root", () => {
    const inventory = propertyFiles();
    const actual = actualPropertyDrivers();
    for (const root of ROOT_TEST_DIRS) {
      const rootFiles = actual.filter((file) => file.startsWith(`${root}/`));
      expect(rootFiles.length).toBeGreaterThan(0);
      expect(
        selectPropertyFiles([`${root}/source.ts`], inventory).map(
          ({ file }) => file,
        ),
      ).toEqual(rootFiles);
      for (const repoPath of rootFiles) {
        expect(
          selectPropertyFiles([repoPath], inventory).map(({ file }) => file),
        ).toEqual([repoPath]);
      }
    }
  });
});

describe("property scheduling profiles", () => {
  const profileSource = (text: string) =>
    propertyFileProfile({ file: "probe.test.ts", text });
  test("sums per-test budgets and applies the file default to unbudgeted drivers", () => {
    const profile = profileSource(`
      const BASE_BUDGET = 30_000;
      setDefaultTimeout(propertyTestTimeout(BASE_BUDGET));
      test("default", () => { assertProperty(p); fc.assert(q); });
      test("explicit", () => { fc.check(r); }, propertyTestTimeout(10_000));
    `);
    expect(profile).toEqual({
      drivers: 3,
      weightMs: 70_000,
      missingBudgets: 0,
    });
  });

  test("a property that starts children weighs its stated subprocess budget", () => {
    expect(
      profileSource(`
      const RUNS = 12;
      setDefaultTimeout(subprocessTestTimeout());
      test("default", () => { assertProperty(p); });
      test("children", () => { assertProperty(q); }, subprocessTestTimeout(RUNS));
      test("own work", () => { fc.check(r); }, subprocessTestTimeout(1, 30_000));
    `),
    ).toEqual({
      drivers: 3,
      weightMs:
        subprocessBudgetBaseMs() +
        subprocessBudgetBaseMs(12) +
        subprocessBudgetBaseMs(1, 30_000),
      missingBudgets: 0,
    });
  });

  test("comments, strings and plain examples do not become property drivers", () => {
    expect(
      profileSource(`
      // assertProperty(p)
      test("plain", () => { const text = "fc.assert(p)"; });
    `),
    ).toEqual({ drivers: 0, weightMs: 0, missingBudgets: 0 });
  });

  test("missing budgets use the documented Bun default estimate without altering execution", () => {
    expect(profileSource("test('p', () => assertProperty(p));")).toEqual({
      drivers: 1,
      weightMs: 5_000,
      missingBudgets: 0,
    });
    expect(
      profileSource(
        "test('p', () => assertProperty(p), propertyTestTimeout());",
      ),
    ).toEqual({
      drivers: 1,
      weightMs: 5_000,
      missingBudgets: 1,
    });
  });

  test("unknown, non-positive and non-finite stated budgets refuse profiling", () => {
    for (const expression of [
      "unknownBudget",
      "0",
      "-1",
      "1e999",
      "computeBudget()",
    ])
      expect(() =>
        profileSource(
          `test('p', () => assertProperty(p), propertyTestTimeout(${expression}));`,
        ),
      ).toThrow();
    expect(() =>
      profileSource(`
      const FIRST = SECOND;
      const SECOND = FIRST;
      test('p', () => assertProperty(p), propertyTestTimeout(FIRST));
    `),
    ).toThrow();
  });
});

describe("property batch execution", () => {
  test("partitioning retains every selected file exactly once in its owning package", () => {
    for (const coreCount of [0, 1, 2, 3, 5]) {
      for (const docxCount of [0, 1, 2, 3, 5]) {
        for (const scriptCount of [0, 1, 2, 3, 5]) {
          const selected = [
            ...Array.from({ length: coreCount }, (_, index) =>
              makePropertyFile(
                `packages/core/src/docx/example-${index}.property.test.ts`,
              ),
            ),
            ...Array.from({ length: docxCount }, (_, index) =>
              makePropertyFile(
                `packages/docx-core/src/model/example-${index}.property.test.ts`,
              ),
            ),
            ...Array.from({ length: scriptCount }, (_, index) => ({
              area: "scripts",
              packageDir: ".",
              file: `scripts/example-${index}.property.test.ts`,
              weightMs: 5_000,
            })),
          ];
          const batches = propertyBatches(selected);
          const actual = batches.flatMap(({ files }) => files);
          expect(actual.map(({ file }) => file).toSorted()).toEqual(
            selected.map(({ file }) => file).toSorted(),
          );
          expect(new Set(actual.map(({ file }) => file)).size).toBe(
            selected.length,
          );
          for (const batch of batches) {
            expect(batch.files.length).toBeGreaterThan(0);
            expect(
              batch.files.every(
                ({ packageDir }) => packageDir === batch.packageDir,
              ),
            ).toBe(true);
          }
          for (const packageDir of new Set(
            selected.map((file) => file.packageDir),
          )) {
            const owned = batches.filter(
              (batch) => batch.packageDir === packageDir,
            );
            expect(owned.length).toBeLessThanOrEqual(2);
            const sizes = owned.map(({ files }) => files.length);
            expect(Math.max(...sizes) - Math.min(...sizes)).toBeLessThanOrEqual(
              1,
            );
          }
          if (selected.length === 0) expect(batches).toEqual([]);
        }
      }
    }
  });

  test("an empty schedule never invokes the runner", async () => {
    let calls = 0;
    expect(
      await runPropertyBatches([], async () => {
        calls += 1;
        return 0;
      }),
    ).toEqual([]);
    expect(calls).toBe(0);
  });

  test("two workers drain every batch and retain failures while another batch remains blocked", async () => {
    const batches = propertyBatches([
      ...ALL,
      makePropertyFile("packages/docx-core/src/model/extra.property.test.ts"),
      {
        area: "scripts",
        packageDir: ".",
        file: "scripts/extra.property.test.ts",
        weightMs: 5_000,
      },
    ]);
    const releases = new Map<
      (typeof batches)[number],
      (code: number) => void
    >();
    const active = new Set<(typeof batches)[number]>();
    const completed = new Set<(typeof batches)[number]>();
    let peak = 0;
    const issuedCodes: number[] = [];
    const execution = runPropertyBatches(batches, (batch) => {
      expect(releases.has(batch)).toBe(false);
      active.add(batch);
      peak = Math.max(peak, active.size);
      expect(active.size).toBeLessThanOrEqual(2);
      return new Promise<number>((resolve) => {
        releases.set(batch, (code) => {
          expect(completed.has(batch)).toBe(false);
          completed.add(batch);
          issuedCodes.push(code);
          active.delete(batch);
          resolve(code);
        });
      });
    });
    await Promise.resolve();
    expect(releases.size).toBe(2);
    expect(active.size).toBe(2);
    const first = releases.entries().next().value;
    if (first === undefined)
      throw new TypeError("Expected the first scheduled batch");
    first[1](17);
    await Promise.resolve();
    await Promise.resolve();
    // A completed worker picks up more work while the other first-wave batch
    // is still blocked. This catches both serial execution and unbounded starts.
    expect(releases.size).toBe(3);
    expect(active.size).toBe(2);
    for (let step = 0; step < batches.length; step += 1) {
      for (const [batch, release] of releases) {
        if (!completed.has(batch))
          release(batches.indexOf(batch) === 2 ? 9 : 0);
      }
      await Promise.resolve();
      await Promise.resolve();
    }
    expect(completed.size).toBe(batches.length);
    expect(releases.size).toBe(batches.length);
    expect(active.size).toBe(0);
    expect(peak).toBe(2);
    const codes = await execution;
    expect(codes).toHaveLength(batches.length);
    expect(codes).toContain(17);
    expect(codes).toContain(9);
    expect(codes.toSorted((left, right) => left - right)).toEqual(
      issuedCodes.toSorted((left, right) => left - right),
    );
  });
});

describe("extended property shards", () => {
  const MINUTE_MS = 60_000;
  // A shard's job also checks out, installs and starts a test process per
  // batch, and runners vary, so its plan keeps a fifth of the timeout free.
  const TIMEOUT_SHARE = 0.8;
  const workflow = Bun.YAML.parse(
    readFileSync(
      path.resolve(import.meta.dir, "../.github/workflows/property-proofs.yml"),
      "utf8",
    ),
  );
  const job = workflow.jobs["property-test"];
  const factor = Number(workflow.on.workflow_dispatch.inputs.factor.default);
  const overLimit = (loads: readonly number[]) =>
    loads.flatMap((load, index) =>
      load > job["timeout-minutes"] * MINUTE_MS * TIMEOUT_SHARE
        ? [
            `shard ${String(index + 1)} plans ${String(load)} ms, over ${String(TIMEOUT_SHARE * 100)}% of ${String(job["timeout-minutes"])} min`,
          ]
        : [],
    );

  test("every planned shard fits within its job's timeout", () => {
    expect(factor).toBe(10);
    const { shards, loads } = planExtendedPropertyShards(
      propertyFiles(),
      factor,
    );
    expect(shards.length).toBeGreaterThan(1);
    expect(overLimit(loads)).toEqual([]);
    expect(overLimit([EXTENDED_PROPERTY_SHARD_BUDGET_MS])).toEqual([]);
  });

  test("a shard planned past the timeout share is reported", () => {
    const limit = job["timeout-minutes"] * MINUTE_MS * TIMEOUT_SHARE;
    expect(overLimit([limit, limit + 1])).toEqual([
      `shard 2 plans ${String(limit + 1)} ms, over 80% of ${String(job["timeout-minutes"])} min`,
    ]);
    // A measured file too long for one shard fails the plan, never runs late.
    const heavy = makePropertyFile("packages/core/src/docx/heavy.test.ts");
    expect(() =>
      planExtendedPropertyShards([heavy], factor, {
        [heavy.file]: Math.ceil(limit / factor) + 1,
      }),
    ).toThrow("on their own");
  });

  test("every property file lands in exactly one shard, as the job selects it", () => {
    const files = propertyFiles();
    const { shards } = planExtendedPropertyShards(files, factor);
    const planned = shards.flatMap((shard) => [...shard]);
    expect(planned.toSorted()).toEqual(
      files.map(({ file }) => file).toSorted(),
    );
    expect(new Set(planned).size).toBe(files.length);
    for (const [position, shard] of shards.entries())
      expect(
        shardPropertyFiles(
          files.toReversed(),
          { index: position + 1, total: shards.length },
          factor,
        ).map(({ file }) => file),
      ).toEqual([...shard].toSorted((a, b) => a.localeCompare(b)));
  });

  test("measured files plan their per-factor time; unmeasured ones their budgets, capped at one shard", () => {
    const measured = makePropertyFile(
      "packages/core/src/docx/m.test.ts",
      1_000,
    );
    const stated = makePropertyFile("packages/core/src/docx/s.test.ts", 2_000);
    const huge = makePropertyFile(
      "packages/core/src/docx/h.test.ts",
      EXTENDED_PROPERTY_SHARD_BUDGET_MS,
    );
    expect(
      propertyShardTimings([measured, stated, huge], factor, {
        [measured.file]: 7,
      }),
    ).toEqual({
      timings: {
        [measured.file]: 7 * factor,
        [stated.file]: 2_000 * factor,
        [huge.file]: EXTENDED_PROPERTY_SHARD_BUDGET_MS,
      },
      unmeasured: [stated.file, huge.file],
    });
  });

  test("the matrix follows the plan and every shard's failures reach the report", () => {
    const plan = workflow.jobs["property-plan"];
    expect(
      plan.steps.find((step: { id?: string }) => step.id === "plan").run,
    ).toBe("bun scripts/property-areas.ts --all --plan");
    expect(plan.outputs.property_shards).toBe(
      "${{ steps.plan.outputs.property_shards }}",
    );
    expect(job.needs).toBe("property-plan");
    expect(job.strategy["fail-fast"]).toBe(false);
    expect(job.strategy.matrix.shard).toBe(
      "${{ fromJSON(needs.property-plan.outputs.property_shards) }}",
    );
    const suite = job.steps.find(
      (step: { name?: string }) => step.name === "Property test suite",
    );
    expect(suite.run).toContain(
      '--all --shard "${PROPERTY_SHARD}/${PROPERTY_SHARD_COUNT}"',
    );
    expect(suite.run).toContain(
      'tee test-results/property-test-"${PROPERTY_SHARD}".log',
    );
    expect(suite.env.PROPERTY_SHARD_COUNT).toBe("${{ strategy.job-total }}");
    const upload = job.steps.find(
      (step: { name?: string }) =>
        step.name === "Keep the log for the failure report",
    );
    expect(upload.with.name).toBe("property-test-log-${{ matrix.shard }}");
    expect(upload.with.path).toBe(
      "test-results/property-test-${{ matrix.shard }}.log",
    );
    const report = workflow.jobs["report-property"];
    expect(report.needs).toBe("property-test");
    const download = report.steps.find(
      (step: { uses?: string }) =>
        step.uses?.startsWith("actions/download-artifact@") === true,
    );
    expect(download.with.pattern).toBe("property-test-log-*");
    expect(download.with["merge-multiple"]).toBe(true);
    expect(JSON.stringify(report.steps)).toContain(
      "property-test-logs/property-test-*.log",
    );
    expect(JSON.stringify(report.steps)).toContain("--log property-test.log");
  });
});
