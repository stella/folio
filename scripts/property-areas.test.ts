import { seedFileFor } from "../test/seed-registry";
import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

import {
  areaOf,
  propertyBatches,
  propertyFiles,
  runPropertyBatches,
  selectPropertyFiles,
  shardPropertyFiles,
  touchedAreas,
} from "./property-areas";

const makePropertyFile = (repoPath: string) => ({
  area: areaOf(repoPath) as string,
  packageDir: repoPath.split("/").slice(0, 2).join("/"),
  file: repoPath,
});

const ALL = [
  makePropertyFile("packages/core/src/ai-edits/batchOverlap.test.ts"),
  makePropertyFile("packages/core/src/compare/compare.property.test.ts"),
  makePropertyFile("packages/core/src/docx/xmlSerialize.property.test.ts"),
  makePropertyFile("packages/docx-core/src/markdown/markdown.test.ts"),
];

/** Independently inventory actual calls, excluding matches in comments and strings. */
const actualPropertyDrivers = (): string[] => {
  const root = path.resolve(import.meta.dir, "..");
  const files = ["packages", "scripts", "test", "parity"].flatMap((dir) =>
    ts.sys.readDirectory(path.join(root, dir), [".ts", ".tsx"], ["**/node_modules/**"]),
  );
  return files
    .filter((file) => /\.test\.tsx?$/.test(file))
    .filter((file) => {
      const source = ts.createSourceFile(
        file,
        readFileSync(file, "utf8"),
        ts.ScriptTarget.Latest,
        true,
      );
      let drives = false;
      const visit = (node: ts.Node): void => {
        if (ts.isCallExpression(node)) {
          const callee = node.expression;
          drives ||=
            (ts.isIdentifier(callee) && callee.text === "assertProperty") ||
            (ts.isPropertyAccessExpression(callee) &&
              ts.isIdentifier(callee.expression) &&
              callee.expression.text === "fc" &&
              (callee.name.text === "assert" || callee.name.text === "check"));
        }
        if (!drives) ts.forEachChild(node, visit);
      };
      visit(source);
      return drives;
    })
    .map((file) => path.relative(root, file).replaceAll("\\", "/"))
    .toSorted((a, b) => a.localeCompare(b));
};

describe("property areas", () => {
  test("the workflow matrix exercises every selected file and collects every failure", () => {
    const workflow = Bun.YAML.parse(
      readFileSync(path.resolve(import.meta.dir, "../.github/workflows/ci.yml"), "utf8"),
    );
    const job = workflow.jobs["property-areas"];
    const files = propertyFiles();
    const exercised = [...job.strategy.matrix.shard].flatMap((index: number) =>
      shardPropertyFiles(files, { index, total: job.strategy.matrix.shard.length }),
    );
    expect(exercised.length).toBe(files.length);
    expect(new Set(exercised).size).toBe(files.length);
    expect(exercised.toSorted((a, b) => a.file.localeCompare(b.file))).toEqual(files);
    expect(job.strategy["fail-fast"]).toBe(false);
    const step = job.steps.find(
      (entry: { name?: string }) => entry.name === "Changed areas' properties at 5x numRuns",
    );
    expect(step.run).toContain('--factor 5 --shard "${PROPERTY_SHARD}/${PROPERTY_SHARD_COUNT}"');
    expect(step.env.PROPERTY_SHARD).toBe("${{ matrix.shard }}");
  });

  test("shards cover every selected file exactly once regardless of input order", () => {
    const inventory = propertyFiles();
    for (const total of [1, 2, 4, inventory.length + 1]) {
      const exercised: string[] = [];
      for (let index = 1; index <= total; index += 1) {
        const shard = { index, total };
        const files = shardPropertyFiles(inventory, shard);
        expect(shardPropertyFiles(inventory.toReversed(), shard)).toEqual(files);
        for (const { file } of files) exercised.push(file);
      }
      expect(exercised.toSorted()).toEqual(inventory.map(({ file }) => file).toSorted());
      expect(new Set(exercised).size).toBe(exercised.length);
    }
    expect(shardPropertyFiles([], { index: 1, total: 4 })).toEqual([]);
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
    expect(areaOf("packages/core/src/compare/diff/align.ts")).toBe("core/compare");
    expect(areaOf("packages/docx-core/src/index.ts")).toBe("docx-core");
    expect(areaOf("scripts/property-areas.ts")).toBeUndefined();
  });

  test("changed source selects its area and the areas coupled to it", () => {
    expect([...touchedAreas(["packages/core/src/internal/compare/blocks.ts"])].toSorted()).toEqual([
      "core/ai-edits",
      "core/compare",
      "core/internal",
      "core/prosemirror",
    ]);
    expect(
      selectPropertyFiles(["packages/docx-core/src/markdown/compile.ts"], ALL).map((f) => f.file),
    ).toEqual(["packages/docx-core/src/markdown/markdown.test.ts"]);
  });

  test("a changed test file reruns itself, not its area", () => {
    expect(
      selectPropertyFiles(
        ["packages/core/src/compare/compare.property.test.ts", "packages/core/src/docx/a.test.ts"],
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
    expect(found.some(({ file: repoPath }) => repoPath.startsWith("packages/core/src/"))).toBe(
      true,
    );
    expect(found.some(({ file: repoPath }) => repoPath.startsWith("packages/docx-core/src/"))).toBe(
      true,
    );
    expect(found.every(({ area }) => area !== undefined)).toBe(true);
    expect(
      found.some(
        ({ file: repoPath }) => repoPath === "scripts/container-survival.property.test.ts",
      ),
    ).toBe(true);
    expect(found.map(({ file }) => file)).toEqual(actualPropertyDrivers());
  });

  test("a root script change selects its properties", () => {
    const rootProperty = {
      area: "scripts",
      packageDir: ".",
      file: "scripts/container-survival.property.test.ts",
    };
    expect(selectPropertyFiles(["scripts/container-survival-census.ts"], [rootProperty])).toEqual([
      rootProperty,
    ]);
  });
});

describe("property batch execution", () => {
  test("partitioning retains every selected file exactly once in its owning package", () => {
    for (const coreCount of [0, 1, 2, 3, 5]) {
      for (const docxCount of [0, 1, 2, 3, 5]) {
        for (const scriptCount of [0, 1, 2, 3, 5]) {
          const selected = [
            ...Array.from({ length: coreCount }, (_, index) =>
              makePropertyFile(`packages/core/src/docx/example-${index}.property.test.ts`),
            ),
            ...Array.from({ length: docxCount }, (_, index) =>
              makePropertyFile(`packages/docx-core/src/model/example-${index}.property.test.ts`),
            ),
            ...Array.from({ length: scriptCount }, (_, index) => ({
              area: "scripts",
              packageDir: ".",
              file: `scripts/example-${index}.property.test.ts`,
            })),
          ];
          const batches = propertyBatches(selected);
          const actual = batches.flatMap(({ files }) => files);
          expect(actual.map(({ file }) => file).toSorted()).toEqual(
            selected.map(({ file }) => file).toSorted(),
          );
          expect(new Set(actual.map(({ file }) => file)).size).toBe(selected.length);
          for (const batch of batches) {
            expect(batch.files.length).toBeGreaterThan(0);
            expect(batch.files.every(({ packageDir }) => packageDir === batch.packageDir)).toBe(
              true,
            );
          }
          for (const packageDir of new Set(selected.map((file) => file.packageDir))) {
            const owned = batches.filter((batch) => batch.packageDir === packageDir);
            expect(owned.length).toBeLessThanOrEqual(2);
            const sizes = owned.map(({ files }) => files.length);
            expect(Math.max(...sizes) - Math.min(...sizes)).toBeLessThanOrEqual(1);
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
      { area: "scripts", packageDir: ".", file: "scripts/extra.property.test.ts" },
    ]);
    const releases = new Map<(typeof batches)[number], (code: number) => void>();
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
    if (first === undefined) throw new TypeError("Expected the first scheduled batch");
    first[1](17);
    await Promise.resolve();
    await Promise.resolve();
    // A completed worker picks up more work while the other first-wave batch
    // is still blocked. This catches both serial execution and unbounded starts.
    expect(releases.size).toBe(3);
    expect(active.size).toBe(2);
    for (let step = 0; step < batches.length; step += 1) {
      for (const [batch, release] of releases) {
        if (!completed.has(batch)) release(batches.indexOf(batch) === 2 ? 9 : 0);
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
