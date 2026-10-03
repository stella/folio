// Binds the CI plan's areas to the jobs they gate. The planner (the shared
// stella/.github ci-plan action) selects areas from `.github/ci-plan.json`;
// `ci.yml` maps each area to a `ci-plan` output and gates a job on it. An area
// no job reads, an output no area feeds, or a job that runs without asking the
// plan would each drift in silence, so all three are checked here.
import { describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import editorWebPlaywright from "../packages/editor-web/playwright.config";
import runContract from "./ci-run-contract.json";
import { PUBLISHED_PACKAGES } from "./lib/published-packages";
import {
  assignTestShards,
  changedTestPaths,
  FAST_SIBLING_TEST_CAP,
  selectFastTestSuites,
  discoverTestSuites,
  focusedTestFiles,
  TEST_SHARD_COUNT,
} from "./ci-test-shards";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PLAN_JOB = "ci-plan";
// The changeset gate runs beside the plan: it has its own pull-request-only
// condition and nothing to scope.
const UNPLANNED_JOBS = new Set([PLAN_JOB, "changeset", "ci-result"]);
const AREA_OUTPUT =
  /^\$\{\{ github\.event_name == 'merge_group' \|\| fromJSON\(steps\.plan\.outputs\.areas\)\.([a-z][a-z0-9_]*) \}\}$/u;
const GATE =
  /^needs\.ci-plan\.outputs\.([a-z][a-z0-9_]*_required) == '(true|false)'(?: && needs\.ci-plan\.outputs\.suite_depth == 'full')?$/u;
const PROPERTY_AREAS_CONDITION =
  "github.event_name == 'pull_request' && github.event.pull_request.draft != true && ";

type Job = { needs?: unknown; if?: unknown; steps?: unknown; outputs?: Record<string, unknown> };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const readWorkflow = (file: string) => {
  const workflow: unknown = Bun.YAML.parse(
    readFileSync(path.join(REPO_ROOT, ".github/workflows", file), "utf8"),
  );
  if (!isRecord(workflow)) throw new Error(`${file} is not a mapping`);
  return workflow;
};

const readJobs = (): Record<string, Job> => {
  const workflow = readWorkflow("ci.yml");
  if (!isRecord(workflow["jobs"])) throw new Error("ci.yml has no jobs");
  const jobs: Record<string, Job> = {};
  for (const [id, job] of Object.entries(workflow["jobs"])) {
    if (!isRecord(job)) throw new Error(`ci.yml job ${id} is not a mapping`);
    const outputs = job["outputs"];
    jobs[id] = {
      needs: job["needs"],
      if: job["if"],
      steps: job["steps"],
      outputs: isRecord(outputs) ? outputs : {},
    };
  }
  return jobs;
};

const readPolicy = () => {
  const policy: unknown = JSON.parse(
    readFileSync(path.join(REPO_ROOT, ".github/ci-plan.json"), "utf8"),
  );
  if (!isRecord(policy) || !isRecord(policy["areas"])) throw new Error("ci-plan.json has no areas");
  return { areas: Object.keys(policy["areas"]), fullDepth: policy["fullDepth"] };
};

describe("CI plan", () => {
  const jobs = readJobs();
  const { areas, fullDepth } = readPolicy();
  const planOutputs = jobs[PLAN_JOB]?.outputs ?? {};

  test("discovery covers every committed Bun test outside declared Playwright directories", () => {
    const tracked = Bun.spawnSync(
      ["git", "ls-files", "--", "packages", "scripts", "benchmarks/compare"],
      {
        cwd: REPO_ROOT,
        stdout: "pipe",
      },
    );
    expect(tracked.exitCode).toBe(0);
    const testDir = editorWebPlaywright.testDir;
    if (typeof testDir !== "string") throw new Error("Editor web Playwright directory is missing");
    const playwrightDirectory = `${path.posix.join("packages/editor-web", testDir)}/`;
    const expected = tracked.stdout
      .toString()
      .trim()
      .split("\n")
      .filter(
        (file) =>
          /(?:\.test|\.spec|_test|_spec)\.(?:js|jsx|ts|tsx|mjs|cjs|mts|cts)$/u.test(file) &&
          !file.startsWith(playwrightDirectory),
      );
    const discovered = discoverTestSuites().flatMap(({ files }) => files);
    expect(discovered.toSorted()).toEqual(expected.toSorted());
  });

  test("each split job retains its complete command list and step depth guards", () => {
    const splitJobs = Object.entries(jobs)
      .filter(
        ([, job]) =>
          job.if === "needs.ci-plan.outputs.code_required == 'true'" ||
          job.if ===
            "needs.ci-plan.outputs.code_required == 'true' && needs.ci-plan.outputs.suite_depth == 'full'",
      )
      .map(([id]) => id);
    expect(Object.keys(runContract).toSorted()).toEqual(splitJobs.toSorted());
    for (const [id, expected] of Object.entries(runContract)) {
      const steps = jobs[id]?.steps;
      if (!Array.isArray(steps)) throw new Error(`Missing steps for ${id}`);
      const commands = steps.flatMap((step) => {
        if (!isRecord(step)) throw new Error(`Invalid step for ${id}`);
        return typeof step["run"] === "string"
          ? [{ run: step["run"], if: step["if"] ?? null }]
          : [];
      });
      expect([id, commands]).toEqual([id, expected]);
    }
    const root = JSON.parse(readFileSync(path.join(REPO_ROOT, "package.json"), "utf8"));
    // Earlier validators leave sibling dist outputs for each later package build.
    expect(root.scripts["validate-dist"]).toBe(
      PUBLISHED_PACKAGES.map(({ slug }) => `bun scripts/validate-dist.ts ${slug}`).join(" && "),
    );
  });

  test("PR source guards retain their required scans without a depth condition", () => {
    const root = JSON.parse(readFileSync(path.join(REPO_ROOT, "package.json"), "utf8"));
    const command = root.scripts["test:source-guards"];
    expect(command.startsWith("bun test ")).toBe(true);
    const guards = command.slice("bun test ".length).split(/\s+/u);
    const required = [
      "scripts/property-test-budgets.test.ts",
      "scripts/rust-boundaries.test.ts",
      "scripts/on-off-element-writer.test.ts",
      "scripts/on-off-spelling.test.ts",
      "scripts/consumer-scenario-dependencies.test.ts",
      "scripts/adapter-layout-timing.test.ts",
      "scripts/specification-sources.test.ts",
      "scripts/ci-plan.test.ts",
    ];
    expect(new Set(guards).size).toBe(guards.length);
    const discovered = new Set(discoverTestSuites().flatMap(({ files }) => files));
    for (const file of guards) expect(discovered.has(file)).toBe(true);
    for (const file of required) expect(guards).toContain(file);
    const steps = jobs["lint"]?.steps;
    if (!Array.isArray(steps)) throw new TypeError("Missing source guard steps");
    const guardSteps = steps.filter(
      (step) => isRecord(step) && step["run"] === "bun run test:source-guards",
    );
    expect(guardSteps).toEqual([
      { name: "Static source guards", run: "bun run test:source-guards" },
    ]);
    expect(jobs["lint"]?.if).toBe("needs.ci-plan.outputs.code_required == 'true'");
  });

  test("discovery includes new nested test files under every command root", () => {
    const fixtureRoot = mkdtempSync(path.join(tmpdir(), "folio-ci-shards-"));
    try {
      const extensions = ["js", "jsx", "ts", "tsx", "mjs", "cjs", "mts", "cts"];
      const stems = ["unit.test", "flow.spec", "unit_test", "flow_spec"];
      const files = stems.flatMap((stem) =>
        extensions.map((extension) => `packages/example/src/nested/${stem}.${extension}`),
      );
      files.push(
        "packages/example/scripts/generator.test.ts",
        "scripts/ci-new.test.mjs",
        "benchmarks/compare/new.spec.ts",
      );
      for (const file of [...files, "packages/example/src/nested/helper.ts"]) {
        const target = path.join(fixtureRoot, file);
        mkdirSync(path.dirname(target), { recursive: true });
        writeFileSync(target, "");
      }
      writeFileSync(
        path.join(fixtureRoot, "packages/example/package.json"),
        JSON.stringify({ scripts: { test: "bun test --preload ./test/setup.ts src scripts" } }),
      );
      const suites = discoverTestSuites(fixtureRoot);
      const discovered = suites.flatMap(({ files: suiteFiles }) => suiteFiles);
      expect(discovered.toSorted()).toEqual(files.toSorted());
      expect(
        assignTestShards(discovered)
          .flatMap((shard) => Array.from(shard))
          .toSorted(),
      ).toEqual(files.toSorted());
      expect(suites.find(({ cwd }) => cwd === "packages/example")?.preloads).toEqual([
        "./test/setup.ts",
      ]);
      for (const command of [
        "bun run custom-tests",
        "bun test src && bun test other",
        "bun test --timeout 5000 src",
      ]) {
        writeFileSync(
          path.join(fixtureRoot, "packages/example/package.json"),
          JSON.stringify({ scripts: { test: command } }),
        );
        expect(() => discoverTestSuites(fixtureRoot)).toThrow("Unsupported test");
      }
    } finally {
      rmSync(fixtureRoot, { recursive: true, force: true });
    }
  });

  test("fast selection includes changed examples and immediate source siblings with preloads", () => {
    const suites = [
      {
        cwd: "packages/example",
        preloads: ["./test/setup.ts"],
        files: [
          "packages/example/src/parser/plain.test.ts",
          "packages/example/src/parser/roundtrip.spec.ts",
          "packages/example/src/parser/nested/unrelated.test.ts",
          "packages/example/src/other/added.test.ts",
          "packages/example/src/other/unchanged.test.ts",
        ],
      },
      { cwd: ".", preloads: [], files: ["scripts/new.test.ts", "scripts/unchanged.test.ts"] },
    ];
    const selected = selectFastTestSuites({
      suites,
      changedFiles: [
        "packages/example/src/parser/parse.ts",
        "packages/example/src/parser/parse.ts",
        "packages/example/src/other/added.test.ts",
        "scripts/new.test.ts",
        "packages/example/src/removed/deleted.test.ts",
      ],
      focusedFiles: [
        "packages/example/src/parser/plain.test.ts",
        "packages/example/src/removed/deleted.test.ts",
      ],
    });
    expect(selected).toEqual([
      {
        cwd: "packages/example",
        preloads: ["./test/setup.ts"],
        files: [
          "packages/example/src/parser/roundtrip.spec.ts",
          "packages/example/src/other/added.test.ts",
        ],
      },
      { cwd: ".", preloads: [], files: ["scripts/new.test.ts"] },
      { cwd: ".", preloads: [], files: ["packages/example/src/parser/plain.test.ts"] },
    ]);
    const selectedFiles = selected.flatMap(({ files }) => files);
    expect(new Set(selectedFiles).size).toBe(selectedFiles.length);
    expect(
      assignTestShards(selectedFiles)
        .flatMap((shard) => Array.from(shard))
        .toSorted(),
    ).toEqual(selectedFiles.toSorted());
    expect(suites.at(0)?.files).toHaveLength(5);
  });

  test("only non-test source changes select sibling tests", () => {
    const suites = [
      {
        cwd: "packages/example",
        preloads: [],
        files: ["packages/example/src/one.test.ts", "packages/example/src/two.test.ts"],
      },
    ];
    expect(
      selectFastTestSuites({
        suites,
        changedFiles: ["packages/example/src/one.test.ts"],
        focusedFiles: [],
      }),
    ).toEqual([{ ...suites[0], files: ["packages/example/src/one.test.ts"] }]);
    expect(
      selectFastTestSuites({
        suites,
        changedFiles: ["packages/example/src/README.md"],
        focusedFiles: [],
      }),
    ).toEqual([]);
  });

  test("the sibling cap expands the owning package without truncation or other packages", () => {
    const siblings = Array.from(
      { length: FAST_SIBLING_TEST_CAP },
      (_, i) => `packages/example/src/parser/test-${i}.test.ts`,
    );
    const elsewhere = "packages/example/src/elsewhere/roundtrip.test.ts";
    const other = {
      cwd: "packages/other",
      preloads: [],
      files: ["packages/other/src/other.test.ts"],
    };
    const select = (files: string[]) =>
      selectFastTestSuites({
        suites: [{ cwd: "packages/example", preloads: ["./setup.ts"], files }, other],
        changedFiles: ["packages/example/src/parser/parse.ts"],
        focusedFiles: [],
      });
    expect(select([...siblings, elsewhere]).at(0)?.files).toEqual(siblings);
    const files = [...siblings, "packages/example/src/parser/over-cap.test.ts", elsewhere];
    expect(select(files)).toEqual([{ cwd: "packages/example", preloads: ["./setup.ts"], files }]);
  });

  test("changed-path discovery excludes deleted tests and deleted source before sibling expansion", () => {
    const fixtureRoot = mkdtempSync(path.join(tmpdir(), "folio-ci-diff-"));
    const git = (...args: string[]) => {
      const result = Bun.spawnSync(["git", ...args], {
        cwd: fixtureRoot,
        stdout: "pipe",
        stderr: "pipe",
      });
      expect(result.exitCode).toBe(0);
      return result.stdout.toString().trim();
    };
    const write = (file: string, text: string) => {
      const target = path.join(fixtureRoot, file);
      mkdirSync(path.dirname(target), { recursive: true });
      writeFileSync(target, text);
    };
    try {
      git("init", "--quiet");
      const deletedSource = "packages/example/src/deleted.ts";
      const deletedTest = "packages/example/src/deleted.test.ts";
      const changedTest = "packages/example/src/changed.test.ts";
      const sibling = "packages/example/src/sibling.test.ts";
      for (const file of [deletedSource, deletedTest, changedTest, sibling]) write(file, "before");
      git("add", ".");
      const commit = () =>
        git(
          "-c",
          "user.name=CI fixture",
          "-c",
          "user.email=ci@example.com",
          "-c",
          "commit.gpgsign=false",
          "commit",
          "--quiet",
          "-m",
          "test: fixture",
        );
      commit();
      const base = git("rev-parse", "HEAD");
      rmSync(path.join(fixtureRoot, deletedSource));
      rmSync(path.join(fixtureRoot, deletedTest));
      write(changedTest, "after");
      const added = "scripts/added.test.ts";
      write(added, "added");
      git("add", ".");
      commit();
      const changedFiles = changedTestPaths(base, fixtureRoot);
      expect(changedFiles.toSorted()).toEqual([changedTest, added].toSorted());
      const selected = selectFastTestSuites({
        suites: [
          { cwd: "packages/example", preloads: [], files: [changedTest, sibling] },
          { cwd: ".", preloads: [], files: [added] },
        ],
        changedFiles,
        focusedFiles: [],
      });
      expect(selected.flatMap(({ files }) => files).toSorted()).toEqual(
        [changedTest, added].toSorted(),
      );
      expect(() => changedTestPaths("missing-base", fixtureRoot)).toThrow(
        "Cannot select changed tests",
      );
    } finally {
      rmSync(fixtureRoot, { recursive: true, force: true });
    }
  });

  test("every discovered and focused test belongs to exactly one deterministic shard", () => {
    const suites = discoverTestSuites();
    const files = [...suites.flatMap(({ files: suiteFiles }) => suiteFiles), ...focusedTestFiles];
    const shards = assignTestShards(files);
    expect(shards).toHaveLength(TEST_SHARD_COUNT);
    const exercised = shards.flatMap((shard) => Array.from(shard));
    expect(exercised.toSorted()).toEqual([...new Set(files)].toSorted());
    expect(shards.map((shard) => Array.from(shard))).toEqual(
      assignTestShards(files.toReversed()).map((shard) => Array.from(shard)),
    );
    // An unmeasured file receives exactly one assignment.
    const future = "packages/core/src/new-area/new-detector.property.test.ts";
    expect(assignTestShards([...files, future]).filter((shard) => shard.has(future))).toHaveLength(
      1,
    );
    expect(suites.find(({ cwd }) => cwd === "packages/vue")?.preloads).toEqual([
      "./test/vueDom.preload.ts",
    ]);
    const root = JSON.parse(readFileSync(path.join(REPO_ROOT, "package.json"), "utf8"));
    expect(root.workspaces).toEqual(["packages/*"]);
    expect(root.scripts.test).toBe(
      "bun --filter '*' test && bun test scripts && bun test benchmarks/compare",
    );
    for (let shard = 1; shard <= TEST_SHARD_COUNT; shard++) {
      const id = `tests-${shard}`;
      expect(jobs[id]?.if).toBe("needs.ci-plan.outputs.code_required == 'true'");
      const steps = jobs[id]?.steps;
      if (!Array.isArray(steps)) throw new Error(`Missing steps for ${id}`);
      const checkout = steps.find(
        (step) =>
          isRecord(step) &&
          typeof step["uses"] === "string" &&
          step["uses"].startsWith("actions/checkout@"),
      );
      if (!isRecord(checkout) || !isRecord(checkout["with"]))
        throw new Error(`Missing checkout for ${id}`);
      expect(checkout["with"]["fetch-depth"]).toBe(0);
      const fast = steps.find(
        (step) =>
          isRecord(step) &&
          step["run"] === `bun scripts/ci-test-shards.ts --shard ${shard} --depth fast`,
      );
      if (!isRecord(fast) || !isRecord(fast["env"]))
        throw new Error(`Missing fast selection for ${id}`);
      expect(fast["env"]["CI_TEST_BASE"]).toBe(
        "${{ github.event_name == 'pull_request' && github.event.pull_request.base.sha || '' }}",
      );
      const uploads = steps.filter(
        (step) =>
          isRecord(step) &&
          typeof step["uses"] === "string" &&
          step["uses"].startsWith("actions/upload-artifact@"),
      );
      expect(uploads).toHaveLength(1);
      const upload = uploads.at(0);
      if (!isRecord(upload) || !isRecord(upload["with"]))
        throw new Error(`Missing upload for ${id}`);
      expect(upload["with"]["name"]).toBe(`fuzz-log-full-test-suite-${shard}`);
      expect(upload["if"]).toBe(
        "failure() && steps.full-suite.outcome == 'failure' && github.event_name == 'merge_group'",
      );
    }
    expect(Object.keys(jobs).filter((job) => /^tests-\d+$/u.test(job))).toHaveLength(
      TEST_SHARD_COUNT,
    );
  });

  // The shared planner's full-depth output can depend on labels or non-PR
  // events. Neither may promote this workflow's depth or expand its path scopes.
  test("only merge groups select full depth and bypass path scopes", () => {
    expect(planOutputs["suite_depth"]).toBe(
      "${{ github.event_name == 'merge_group' && 'full' || 'fast' }}",
    );
    expect(fullDepth).toBe("scoped");
  });

  test("random browser input fuzz runs only outside PR and merge gates", () => {
    const nightly = readWorkflow("nightly-browser-input-fuzzer.yml");
    const triggers = nightly["on"];
    if (!isRecord(triggers)) throw new Error("browser fuzz workflow is missing triggers");
    expect(Object.keys(triggers).toSorted()).toEqual(["schedule", "workflow_dispatch"]);
    expect(JSON.stringify(readWorkflow("ci.yml"))).not.toContain("--project=browser-fuzzer");
    expect(JSON.stringify(nightly)).toContain("--project=browser-fuzzer");
    expect(jobs).not.toHaveProperty("browser-fuzzer-smoke");
  });

  // CodSpeed rejects merge_group events, so benchmarks keep their own
  // pull-request, main-push and nightly triggers.
  test("benchmarks never run on merge groups", () => {
    const triggers = readWorkflow("benchmarks.yml")["on"];
    if (!isRecord(triggers)) throw new Error("benchmarks.yml is missing triggers");
    expect(Object.keys(triggers)).not.toContain("merge_group");
    expect(Object.keys(triggers)).toContain("push");
  });

  for (const file of ["oracle-mutation-check.yml", "vscode-extension.yml"]) {
    test(`${file} runs heavy jobs only for merge groups`, () => {
      const workflow = readWorkflow(file);
      const triggers = workflow["on"];
      const heavyJobs = workflow["jobs"];
      if (!isRecord(triggers) || !isRecord(heavyJobs)) {
        throw new Error(`${file} is missing triggers or jobs`);
      }
      expect(Object.keys(triggers).toSorted()).toEqual(["merge_group", "pull_request"]);
      expect(triggers["pull_request"]).toEqual({
        types: ["opened", "synchronize", "reopened", "ready_for_review"],
      });
      expect(triggers["merge_group"]).toEqual({ types: ["checks_requested"] });
      expect(Object.keys(heavyJobs).length).toBeGreaterThan(0);
      for (const job of Object.values(heavyJobs)) {
        if (!isRecord(job)) throw new Error(`${file} contains an invalid job`);
        expect(job["if"]).toBe("github.event_name == 'merge_group'");
      }
    });
  }

  test("every area is a plan output named after it, and every area output is an area", () => {
    const mapped = Object.entries(planOutputs).flatMap(([output, value]) => {
      const area = typeof value === "string" ? AREA_OUTPUT.exec(value)?.[1] : undefined;
      return area === undefined ? [] : [{ output, area }];
    });
    expect(mapped.map(({ area }) => area).toSorted()).toEqual(areas.toSorted());
    for (const { output, area } of mapped) expect(output).toBe(`${area}_required`);
  });

  test("every other job waits on the plan and is gated by one area it maps", () => {
    const gated = new Set<string>();
    for (const [id, job] of Object.entries(jobs)) {
      if (UNPLANNED_JOBS.has(id)) continue;
      expect([id, job.needs]).toEqual([id, PLAN_JOB]);
      const condition = job.if;
      if (id === "property-areas") {
        expect(
          typeof condition === "string" && condition.startsWith(PROPERTY_AREAS_CONDITION),
        ).toBe(true);
      }
      const areaGate =
        id === "property-areas" && typeof condition === "string"
          ? condition.slice(PROPERTY_AREAS_CONDITION.length)
          : condition;
      const gate = typeof areaGate === "string" ? GATE.exec(areaGate) : null;
      expect([id, gate !== null]).toEqual([id, true]);
      const output = gate?.[1] ?? "";
      expect([id, output in planOutputs]).toEqual([id, true]);
      if (gate?.[2] === "true") gated.add(output);
    }
    // No area is planned for nothing.
    expect([...gated].toSorted()).toEqual(areas.map((area) => `${area}_required`).toSorted());
  });
});
