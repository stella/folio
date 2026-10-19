import { scanTestTimeout } from "../test/scan-timeout";

// Binds the CI plan's areas to the jobs they gate. The planner (the shared
// stella/.github ci-plan action) selects areas from `.github/ci-plan.json`;
// `ci.yml` maps each area to a `ci-plan` output and gates a job on it. An area
// no job reads, an output no area feeds, or a job that runs without asking the
// plan would each drift in silence, so all three are checked here.
import { setDefaultTimeout, describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import editorWebPlaywright from "../packages/editor-web/playwright.config";
import runContract from "./ci-run-contract.json";
import { affectedPackagePlan } from "./ci-api-plan";
import {
  BROWSER_INPUT_REGRESSIONS,
  PINNED_BROWSER_INPUT_TAG,
} from "../tests/visual/browserInputRegressions";
import {
  assignTestShards,
  changedTestPaths,
  FAST_SIBLING_TEST_CAP,
  selectFastTestSuites,
  selectFullTestSuites,
  discoverTestSuites,
  focusedTestFiles,
  TEST_SHARD_COUNT,
} from "./ci-test-shards";

setDefaultTimeout(scanTestTimeout(30_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PLAN_JOB = "ci-plan";
// The changeset gate runs beside the plan: it has its own pull-request-only
// condition and nothing to scope. Lint checks policy unless a push reuses CI.
const UNPLANNED_JOBS = new Set([PLAN_JOB, "changeset", "lint", "warm-main-caches", "ci-result"]);
const AREA_OUTPUT =
  /^\$\{\{ steps\.reuse\.outputs\.reused_run == '' && \(github\.event_name == 'merge_group' \|\| steps\.toolchain\.outputs\.changed == 'true' \|\| (?:steps\.api-plan\.outputs\.packaged_consumer_required == 'true' \|\| )?fromJSON\(steps\.plan\.outputs\.areas\)\.([a-z][a-z0-9_]*)\) \}\}$/u;
const GATE =
  /^needs\.ci-plan\.outputs\.([a-z][a-z0-9_]*_required) == '(true|false)'(?: && needs\.ci-plan\.outputs\.suite_depth == 'full')?$/u;
const PR_SCOPED_JOBS = new Set(["browser-discovery"]);
const PROPERTY_AREAS_CONDITION =
  "github.event_name == 'pull_request' && github.event.pull_request.draft != true && ";

type Job = {
  needs?: unknown;
  if?: unknown;
  steps?: unknown;
  outputs?: Record<string, unknown>;
  "timeout-minutes"?: unknown;
  strategy?: unknown;
};

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
      "timeout-minutes": job["timeout-minutes"],
      strategy: job["strategy"],
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
        ([id, job]) =>
          id === "lint" ||
          job.if === "needs.ci-plan.outputs.code_required == 'true'" ||
          job.if ===
            "needs.ci-plan.outputs.code_required == 'true' && needs.ci-plan.outputs.suite_depth == 'full'" ||
          job.if === "needs.ci-plan.outputs.api_required == 'true'",
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
      "scripts/zip-part-writer.test.ts",
      "scripts/encoding-boundaries.test.ts",
      "scripts/text-truncation.test.ts",
      "scripts/on-off-spelling.test.ts",
      "scripts/consumer-scenario-dependencies.test.ts",
      "scripts/check-package-dependencies.test.ts",
      "scripts/packed-declaration-diagnostics.test.ts",
      "scripts/packed-exports.test.ts",
      "scripts/published-packages.test.ts",
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
    expect(jobs["lint"]?.if).toBe("needs.ci-plan.outputs.reused_run == ''");
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

  test("full selection covers every discovered and focused file exactly once with owner preloads", () => {
    const suites = discoverTestSuites();
    const discovered = suites.flatMap(({ files }) => files);
    const selected = selectFullTestSuites({ suites });
    const full = selected.flatMap(({ files }) => files);
    expect(full.toSorted()).toEqual([...new Set([...discovered, ...focusedTestFiles])].toSorted());
    expect(new Set(full).size).toBe(full.length);
    for (const suite of suites) {
      expect(selected).toContain(suite);
    }
    for (const changedFiles of [[], discovered, focusedTestFiles]) {
      const fast = selectFastTestSuites({ suites, changedFiles });
      for (const file of fast.flatMap(({ files }) => files)) expect(full).toContain(file);
    }
    expect(selectFullTestSuites({ suites, focusedFiles: [...discovered, ...discovered] })).toBe(
      suites,
    );
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
      expect(fast["if"]).toBe("needs.ci-plan.outputs.suite_depth == 'fast'");
      const full = steps.find((step) => isRecord(step) && step["id"] === "full-suite");
      if (!isRecord(full)) throw new Error(`Missing full selection for ${id}`);
      expect(full["if"]).toBe("needs.ci-plan.outputs.suite_depth == 'full'");
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

  // Path planning stays scoped; a shared toolchain change promotes the whole suite.
  test("merge groups and shared toolchain changes select full depth", () => {
    expect(planOutputs["toolchain_changed"]).toBe("${{ steps.toolchain.outputs.changed }}");
    expect(planOutputs["suite_depth"]).toBe(
      "${{ (github.event_name == 'merge_group' || steps.toolchain.outputs.changed == 'true') && 'full' || 'fast' }}",
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

  test(
    "pinned browser inputs gate PRs and merge groups with exactly the standing registry",
    () => {
      const pinned = jobs["browser-input-pinned"];
      expect(pinned?.if).toBe("needs.ci-plan.outputs.interactions_required == 'true'");
      const steps = pinned?.steps;
      if (!Array.isArray(steps)) throw new TypeError("Missing pinned browser steps");
      const commands = steps.flatMap((step) =>
        isRecord(step) && typeof step["run"] === "string" ? [step["run"]] : [],
      );
      expect(commands.filter((command) => command.includes("playwright test"))).toEqual([
        "bunx playwright test --project=browser-input-pinned --workers=1",
      ]);
      expect(pinned?.strategy).toBeUndefined();
      expect(pinned?.["timeout-minutes"]).toBe(10);

      const listing = Bun.spawnSync(
        [
          "bunx",
          "playwright",
          "test",
          "--project=browser-input-pinned",
          "--list",
          "--reporter=json",
        ],
        { cwd: REPO_ROOT, stdout: "pipe", stderr: "pipe" },
      );
      expect(listing.exitCode, listing.stderr.toString()).toBe(0);
      const report: unknown = JSON.parse(listing.stdout.toString());
      if (!isRecord(report)) throw new TypeError("Invalid pinned project report");
      expect(report["errors"]).toEqual([]);
      const titles: string[] = [];
      const collect = (suites: unknown) => {
        if (!Array.isArray(suites)) throw new TypeError("Invalid discovery suites");
        for (const suite of suites) {
          if (!isRecord(suite) || !Array.isArray(suite["specs"]))
            throw new TypeError("Invalid discovery suite");
          for (const spec of suite["specs"]) {
            if (
              !isRecord(spec) ||
              typeof spec["title"] !== "string" ||
              !Array.isArray(spec["tests"])
            )
              throw new TypeError("Invalid pinned spec");
            expect(spec["tags"]).toEqual([PINNED_BROWSER_INPUT_TAG.slice(1)]);
            for (const discovered of spec["tests"]) {
              if (!isRecord(discovered)) throw new TypeError("Invalid pinned test");
              expect(discovered["projectName"]).toBe("browser-input-pinned");
              expect(discovered["timeout"]).toBe(60_000);
              titles.push(spec["title"]);
            }
          }
          if (suite["suites"] !== undefined) collect(suite["suites"]);
        }
      };
      collect(report["suites"]);
      expect(titles.toSorted()).toEqual(
        BROWSER_INPUT_REGRESSIONS.map(
          ({ seed, path: replayPath, fingerprint }) =>
            `regression seed ${seed} path ${replayPath} / ${fingerprint}`,
        ).toSorted(),
      );
      expect(titles.length).toBeGreaterThan(0);
    },
    scanTestTimeout(30_000),
  );

  test("benchmarks and extended proofs run only on demand", () => {
    for (const file of [
      "benchmarks.yml",
      "oracle-mutation-check.yml",
      "nightly-webkit-parity.yml",
      "property-proofs.yml",
      "parity-e2e.yml",
      "corpus-proofs.yml",
      "differential-proofs.yml",
    ]) {
      const workflow = readWorkflow(file);
      const triggers = workflow["on"];
      if (!isRecord(triggers)) throw new Error(`${file} is missing triggers`);
      expect(Object.keys(triggers)).toEqual(["workflow_dispatch"]);
      const proofJobs = workflow["jobs"];
      if (!isRecord(proofJobs)) throw new Error(`${file} is missing jobs`);
      for (const job of Object.values(proofJobs)) {
        if (!isRecord(job)) throw new Error(`${file} contains an invalid job`);
        expect(job["if"]).not.toBe("github.event_name == 'merge_group'");
        const permissions = job["permissions"] ?? workflow["permissions"];
        if (!isRecord(permissions)) throw new Error(`${file} lacks checkout permissions`);
        expect(permissions["contents"]).toBe("read");
      }
    }
  });

  test("mixed workflows preserve scheduled fuzz and move every proof job", () => {
    const lanes = [
      {
        scheduled: "nightly-property-test.yml",
        fuzz: ["consumer-fuzz", "report-consumer-fuzz"],
        manual: "property-proofs.yml",
        proofs: [
          "property-test",
          "docx-kernel-property",
          "conformance-full",
          "report-property",
          "report-conformance",
          "property-areas",
        ],
      },
      {
        scheduled: "nightly-corpus-gate.yml",
        fuzz: ["edit-fuzz", "edit-fuzz-ratchet"],
        manual: "corpus-proofs.yml",
        proofs: ["shard", "tier2", "ratchet"],
      },
    ];
    for (const lane of lanes) {
      const automatic = readWorkflow(lane.scheduled);
      const manual = readWorkflow(lane.manual);
      if (
        !isRecord(automatic["jobs"]) ||
        !isRecord(manual["jobs"]) ||
        !isRecord(automatic["on"]) ||
        !isRecord(manual["on"])
      )
        throw new Error("Mixed lane mappings unavailable");
      expect(Object.keys(automatic["jobs"])).toEqual(lane.fuzz);
      expect(Object.keys(manual["jobs"])).toEqual(lane.proofs);
      expect(Object.keys(automatic["on"]).toSorted()).toEqual(["schedule", "workflow_dispatch"]);
      expect(Object.keys(manual["on"])).toEqual(["workflow_dispatch"]);
    }
  });

  test("full-corpus comparison and raised property runs stay outside correctness CI", () => {
    expect(jobs).not.toHaveProperty("property-areas");
    expect(JSON.stringify(readWorkflow("ci.yml"))).not.toContain("scripts/property-areas.ts");
    expect(JSON.stringify(readWorkflow("ci.yml"))).not.toContain("bun run test:differential");
    const property = readWorkflow("property-proofs.yml");
    expect(JSON.stringify(property)).toContain("scripts/property-areas.ts --base");
    const differential = readWorkflow("differential-proofs.yml");
    expect(JSON.stringify(differential)).toContain("bun run test:differential");
    expect(JSON.stringify(differential)).toContain('"DIFFERENTIAL_REQUIRED":"1"');
    expect(JSON.stringify(jobs.differential)).toContain(
      "bun test benchmarks/compare/documents.test.ts benchmarks/compare/generated-roundtrip.test.ts",
    );
  });

  const vscodePackagingPaths = (workflow: Record<string, unknown>): string[] => {
    const build = readFileSync(path.join(REPO_ROOT, "plugins/vscode/scripts/build.ts"), "utf8");
    const roots = [...build.matchAll(/path\.join\(repoRoot,\s*"packages",\s*"([^"]+)"/gu)].map(
      (match) => match.at(1),
    );
    expect(roots.length).toBeGreaterThan(0);
    const packages = new Map<string, { slug: string; manifest: Record<string, unknown> }>();
    for (const directory of readdirSync(path.join(REPO_ROOT, "packages"), {
      withFileTypes: true,
    })) {
      if (!directory.isDirectory()) continue;
      const manifest: unknown = JSON.parse(
        readFileSync(path.join(REPO_ROOT, "packages", directory.name, "package.json"), "utf8"),
      );
      if (!isRecord(manifest) || typeof manifest["name"] !== "string")
        throw new Error("Workspace manifest unavailable");
      packages.set(manifest["name"], { slug: directory.name, manifest });
    }
    const selected = new Set<string>();
    const queue = roots.map((slug) => {
      const workspace = [...packages.values()].find((entry) => entry.slug === slug);
      if (!workspace) throw new Error(`Unknown build workspace ${slug}`);
      return workspace;
    });
    for (const workspace of queue) {
      if (selected.has(workspace.slug)) continue;
      selected.add(workspace.slug);
      for (const field of ["dependencies", "peerDependencies"]) {
        const dependencies = workspace.manifest[field];
        if (!isRecord(dependencies)) continue;
        for (const [name, version] of Object.entries(dependencies)) {
          const dependency = packages.get(name);
          if (dependency) queue.push(dependency);
          else if (typeof version === "string" && version.startsWith("workspace:"))
            throw new Error(`Unknown workspace dependency ${name}`);
        }
      }
    }
    const paths = new Set([
      "plugins/vscode/**",
      ".github/workflows/vscode-extension.yml",
      "scripts/ci-plan.test.ts",
    ]);
    for (const slug of selected) paths.add(`packages/${slug}/**`);
    for (const match of build.matchAll(/path\.join\(repoRoot,\s*"([^"]+)"\)/gu)) {
      const file = match.at(1);
      if (!file) throw new Error("Missing build input path");
      paths.add(file);
    }
    const smoke = readFileSync(path.join(REPO_ROOT, "plugins/vscode/test/smoke/run.ts"), "utf8");
    for (const match of smoke.matchAll(/path\.join\(repoRoot,\s*((?:"[^"]+"\s*,?\s*)+)\)/gu)) {
      const argumentsText = match.at(1);
      if (!argumentsText) throw new Error("Missing smoke fixture arguments");
      const parts = [...argumentsText.matchAll(/"([^"]+)"/gu)].map((part) => {
        const value = part.at(1);
        if (!value) throw new Error("Missing smoke fixture path");
        return value;
      });
      paths.add(path.posix.join(...parts));
    }
    const packagingJobs = workflow["jobs"];
    if (
      !isRecord(packagingJobs) ||
      !isRecord(packagingJobs["package"]) ||
      !Array.isArray(packagingJobs["package"]["steps"])
    )
      throw new Error("Missing packaging steps");
    for (const step of packagingJobs["package"]["steps"]) {
      if (!isRecord(step)) throw new Error("Invalid packaging step");
      if (isRecord(step["with"])) {
        for (const key of ["node-version-file", "bun-version-file"]) {
          const file = step["with"][key];
          if (typeof file === "string") paths.add(file);
        }
      }
      if (typeof step["run"] !== "string") continue;
      for (const match of step["run"].matchAll(/\.\.\/\.\.\/([^\s"\\]+\.docx)/gu)) {
        const file = match.at(1);
        if (!file) throw new Error("Missing packaging fixture path");
        paths.add(file);
      }
      if (step["working-directory"] !== undefined) continue;
      if (step["run"].includes("bun install")) {
        paths.add("bun.lock");
        paths.add("bunfig.toml");
      }
      for (const match of step["run"].matchAll(/\bbun (scripts\/[\w./-]+\.ts)/gu)) {
        const file = match.at(1);
        if (!file) throw new Error("Missing workflow script path");
        paths.add(file);
      }
    }
    const config: unknown = JSON.parse(
      readFileSync(path.join(REPO_ROOT, "plugins/vscode/tsconfig.json"), "utf8"),
    );
    if (!isRecord(config) || typeof config["extends"] !== "string")
      throw new Error("Missing extension compiler base");
    paths.add(
      path.relative(REPO_ROOT, path.resolve(REPO_ROOT, "plugins/vscode", config["extends"])),
    );
    return [...paths].toSorted();
  };

  for (const file of ["vscode-extension.yml"]) {
    test(`${file} packages plugin pull requests and all merge groups`, () => {
      const workflow = readWorkflow(file);
      const triggers = workflow["on"];
      const heavyJobs = workflow["jobs"];
      if (!isRecord(triggers) || !isRecord(heavyJobs)) {
        throw new Error(`${file} is missing triggers or jobs`);
      }
      expect(Object.keys(triggers).toSorted()).toEqual(["merge_group", "pull_request"]);
      expect(triggers["pull_request"]).toEqual({
        types: ["opened", "synchronize", "reopened", "ready_for_review"],
        paths: vscodePackagingPaths(workflow),
      });
      expect(triggers["merge_group"]).toEqual({ types: ["checks_requested"] });
      expect(Object.keys(heavyJobs).length).toBeGreaterThan(0);
      for (const job of Object.values(heavyJobs)) {
        if (!isRecord(job)) throw new Error(`${file} contains an invalid job`);
        expect(job["if"]).toBe(
          "github.event_name == 'merge_group' || github.event.pull_request.draft != true",
        );
      }
    });
  }

  test("packed dependency smoke runs on scoped PRs and shares tarballs with full artifact checks", () => {
    expect(jobs["packaged-consumer"]?.if).toBe(
      "needs.ci-plan.outputs.packaged_consumer_required == 'true'",
    );
    expect(planOutputs["packaged_consumer_required"]).toContain(
      "steps.api-plan.outputs.packaged_consumer_required == 'true'",
    );
    const steps = jobs["packaged-consumer"]?.steps;
    if (!Array.isArray(steps)) throw new TypeError("Missing packaged consumer steps");
    const smoke = steps.find(
      (step) => isRecord(step) && step["name"] === "Dependency smoke over every published tarball",
    );
    expect(smoke).toEqual({
      name: "Dependency smoke over every published tarball",
      run: "bun scripts/packaged-consumer-dependencies.ts --pack-dir test-results/package-tarballs",
    });
    for (const name of [
      "Validate published artifacts",
      "Production Vite build over packed tarballs",
      "Runtime smoke over packed tarballs",
    ]) {
      const step = steps.find((candidate) => isRecord(candidate) && candidate["name"] === name);
      if (!isRecord(step)) throw new TypeError(`Missing packed step ${name}`);
      expect(step["if"]).toBe("needs.ci-plan.outputs.suite_depth == 'full'");
      expect(step["env"]).toEqual({ FOLIO_PACKED_TARBALLS: "test-results/package-tarballs" });
    }
    const lintSteps = jobs["lint"]?.steps;
    if (!Array.isArray(lintSteps)) throw new TypeError("Missing lint steps");
    expect(
      lintSteps.find(
        (step) => isRecord(step) && step["name"] === "Published package dependency declarations",
      ),
    ).toEqual({
      name: "Published package dependency declarations",
      run: "bun scripts/check-package-dependencies.ts --source",
    });
    expect(jobs["validate-dist"]).toBeUndefined();
    const harness = readFileSync(
      path.join(REPO_ROOT, "scripts/packaged-consumer-dependencies.ts"),
      "utf8",
    );
    expect(harness).toContain('import { PUBLISHED_PACKAGES } from "./lib/published-packages"');
    // Both packing and consumption must traverse the discovered publication set.
    expect(harness.match(/for \(const pkg of PUBLISHED_PACKAGES\)/gu)).toHaveLength(2);
  });

  test("packed consumer scenarios follow the pull-request package plan", () => {
    const condition = "needs.ci-plan.outputs.packaged_consumer_required == 'true'";
    expect(jobs["consumer-scenarios"]?.if).toBe(condition);
    expect(jobs["consumer-scenarios-minimum"]?.if).toBe(condition);

    const publishedPackagePaths = ["packages/public"];
    for (const changedPath of [
      "packages/public/src/index.ts",
      "test/consumer-scenarios/scenarios/compare.test.ts",
    ]) {
      expect(
        affectedPackagePlan({ changedPaths: [changedPath], publishedPackagePaths })
          .packagedConsumerRequired,
      ).toBe(true);
    }
    expect(
      affectedPackagePlan({ changedPaths: ["docs/guide.md"], publishedPackagePaths })
        .packagedConsumerRequired,
    ).toBe(false);

    const resultSteps = jobs["ci-result"]?.steps;
    if (!Array.isArray(resultSteps)) throw new TypeError("Missing CI result steps");
    const evaluation = resultSteps.find(
      (step) => isRecord(step) && step["name"] === "Evaluate CI outcome",
    );
    if (!isRecord(evaluation) || !isRecord(evaluation["env"]))
      throw new TypeError("Missing CI result environment");
    const scopes = JSON.parse(String(evaluation["env"]["JOB_SCOPES"]));
    expect(scopes["consumer-scenarios"]).toEqual({ area: "packaged_consumer_required" });
    expect(scopes["consumer-scenarios-minimum"]).toEqual({
      area: "packaged_consumer_required",
    });
  });

  test("VS Code compatibility gates its selected area and exercises both API branches", () => {
    const workflow = readWorkflow("vscode-extension.yml");
    const triggers = workflow["on"];
    if (!isRecord(triggers)) throw new Error("Missing VS Code workflow triggers");
    expect(Object.keys(triggers).toSorted()).toEqual(["merge_group", "pull_request"]);
    if (!isRecord(workflow["jobs"]) || !isRecord(workflow["jobs"]["package"]))
      throw new Error("Missing VS Code package job");
    const steps = workflow["jobs"]["package"]["steps"];
    if (!Array.isArray(steps)) throw new Error("Missing VS Code test steps");
    const smoke = steps.filter(
      (step) => isRecord(step) && step["run"] === "xvfb-run -a bun run test:smoke",
    );
    expect(smoke.map((step) => step["env"])).toEqual([
      { FOLIO_SMOKE_VSCODE_VERSION: "1.86.0", FOLIO_SMOKE_MCP: "unavailable" },
      { FOLIO_SMOKE_VSCODE_VERSION: "stable", FOLIO_SMOKE_MCP: "registered" },
    ]);
    expect(
      steps.filter((step) => isRecord(step) && step["run"] === "bun run package"),
    ).toHaveLength(1);
  });

  test("published API preflight is selected before queue entry and evaluated by CI result", () => {
    expect(planOutputs["api_required"]).toBe(
      "${{ steps.reuse.outputs.reused_run == '' && (github.event_name == 'merge_group' || steps.toolchain.outputs.changed == 'true' || steps.api-plan.outputs.api_required == 'true') }}",
    );
    expect(jobs["build-api"]?.if).toBe("needs.ci-plan.outputs.api_required == 'true'");
    const steps = jobs[PLAN_JOB]?.steps;
    if (!Array.isArray(steps)) throw new TypeError("Missing planner steps");
    const selector = steps.find((step) => isRecord(step) && step["id"] === "api-plan");
    expect(selector).toEqual({
      name: "Select published API preflight",
      id: "api-plan",
      if: "github.event_name == 'pull_request'",
      run: 'bun scripts/ci-api-plan.ts --base "$API_BASE" --head "$API_HEAD"',
      env: {
        API_BASE: "${{ github.event.pull_request.base.sha }}",
        API_HEAD: "${{ github.event.pull_request.head.sha }}",
      },
    });
    const resultSteps = jobs["ci-result"]?.steps;
    if (!Array.isArray(resultSteps)) throw new TypeError("Missing CI result steps");
    const evaluation = resultSteps.find(
      (step) => isRecord(step) && step["name"] === "Evaluate CI outcome",
    );
    if (!isRecord(evaluation) || !isRecord(evaluation["env"]))
      throw new TypeError("Missing CI result environment");
    const scopes = JSON.parse(String(evaluation["env"]["JOB_SCOPES"]));
    expect(scopes["build-api"]).toEqual({ area: "api_required" });
  });

  test("browser discovery path scopes include future configs and playgrounds but skip docs", () => {
    const policy: unknown = JSON.parse(
      readFileSync(path.join(REPO_ROOT, ".github/ci-plan.json"), "utf8"),
    );
    if (!isRecord(policy) || !isRecord(policy["areas"])) throw new TypeError("Missing CI areas");
    const scope = policy["areas"]["browser_discovery"];
    if (!isRecord(scope) || !Array.isArray(scope["paths"]))
      throw new TypeError("Missing discovery paths");
    const patterns = scope["paths"].map((pattern) => {
      if (typeof pattern !== "string") throw new TypeError("Invalid discovery pattern");
      return pattern;
    });
    const fixtures = {
      "tests/visual/**": "tests/visual/new.interactions.spec.ts",
      "tests/parity/**": "tests/parity/new.spec.ts",
      "**/playwright*.config.*": "new-package/playwright.future.config.mts",
      "packages/playground*/**": "packages/playground-future/src/main.ts",
      "packages/editor-web/e2e/**": "packages/editor-web/e2e/new.spec.ts",
      "!**/*.md": "tests/visual/README.md",
    };
    expect(patterns.toSorted()).toEqual(Object.keys(fixtures).toSorted());
    const selected = (file: string) =>
      patterns.some((pattern) => !pattern.startsWith("!") && new Bun.Glob(pattern).match(file)) &&
      !patterns.some(
        (pattern) => pattern.startsWith("!") && new Bun.Glob(pattern.slice(1)).match(file),
      );
    for (const [pattern, file] of Object.entries(fixtures))
      expect(selected(file)).toBe(!pattern.startsWith("!"));
    expect(selected("playwright.config.ts")).toBe(true);
    expect(selected("docs/browser-tests.md")).toBe(false);
    expect(selected("packages/docx-core/src/model.ts")).toBe(false);
  });

  test("browser discovery runs on scoped PRs without changing merge-group execution", () => {
    const steps = jobs["browser-discovery"]?.steps;
    if (!Array.isArray(steps)) throw new TypeError("Missing browser discovery steps");
    expect(jobs["browser-discovery"]?.if).toBe(
      `${PROPERTY_AREAS_CONDITION}needs.ci-plan.outputs.browser_discovery_required == 'true'`,
    );
    expect(steps.filter((step) => isRecord(step) && typeof step["run"] === "string")).toEqual([
      { run: "bun scripts/ci-scoped-install.ts browser-discovery" },
      {
        name: "Discover every Node Playwright project without browsers",
        run: "bun scripts/ci-browser-discovery.ts",
      },
    ]);
    const resultSteps = jobs["ci-result"]?.steps;
    if (!Array.isArray(resultSteps)) throw new TypeError("Missing CI result steps");
    const evaluation = resultSteps.find(
      (step) => isRecord(step) && step["name"] === "Evaluate CI outcome",
    );
    if (!isRecord(evaluation) || !isRecord(evaluation["env"]))
      throw new TypeError("Missing CI result environment");
    const scopes = JSON.parse(String(evaluation["env"]["JOB_SCOPES"]));
    expect(scopes["browser-discovery"]).toEqual({
      event: "pull_request",
      area: "browser_discovery_required",
    });
  });

  test("CI outcome evaluation owns every registered job and preserves its gate", () => {
    const result = jobs["ci-result"];
    const steps = result?.steps;
    if (!Array.isArray(steps)) throw new TypeError("Missing CI result steps");
    const evaluation = steps.find(
      (step) => isRecord(step) && step["name"] === "Evaluate CI outcome",
    );
    if (!isRecord(evaluation) || !isRecord(evaluation["env"]))
      throw new TypeError("Missing CI result environment");
    const scopes: unknown = JSON.parse(String(evaluation["env"]["JOB_SCOPES"]));
    if (!isRecord(scopes)) throw new TypeError("Invalid outcome scopes");
    const evaluated = Object.keys(jobs).filter((id) => id !== PLAN_JOB && id !== "ci-result");
    expect(Object.keys(scopes).toSorted()).toEqual(evaluated.toSorted());
    expect(result?.needs).toEqual(expect.arrayContaining([PLAN_JOB, ...evaluated]));
    if (!Array.isArray(result?.needs)) throw new TypeError("Missing CI outcome dependencies");
    expect(result.needs.length).toBe(evaluated.length + 1);
    for (const id of evaluated) {
      const condition = jobs[id]?.if;
      const scope = scopes[id];
      if (!isRecord(scope)) throw new TypeError(`Missing job scope ${id}`);
      if (typeof scope["area"] !== "string") continue;
      expect(typeof condition).toBe("string");
      if (typeof condition !== "string") throw new TypeError(`Missing job gate ${id}`);
      const suffix =
        scope["depth"] === "full" ? " && needs.ci-plan.outputs.suite_depth == 'full'" : "";
      const areaGate = `needs.ci-plan.outputs.${scope["area"]} == '${scope["value"] ?? "true"}'${suffix}`;
      expect(condition.endsWith(areaGate)).toBe(true);
    }
    expect(scopes["lint"]).toEqual({ reuse: "forbidden" });
    expect(scopes["docs-format"]).toEqual({
      area: "code_required",
      value: "false",
      reuse: "forbidden",
    });
    expect(scopes["warm-main-caches"]).toEqual({
      event: "push",
      reuse: "required",
      warm: "required",
    });
    expect(scopes["browser-input-pinned"]).toEqual({ area: "interactions_required" });
  });

  test("every required output disables fresh work when CI is reused", () => {
    const required = Object.entries(planOutputs).filter(([output]) => output.endsWith("_required"));
    expect(required.length).toBeGreaterThan(0);
    for (const [output, value] of required) {
      expect([
        output,
        typeof value === "string" &&
          value.startsWith("${{ steps.reuse.outputs.reused_run == '' && (") &&
          value.endsWith(") }}"),
      ]).toEqual([output, true]);
    }
    expect(jobs["docs-format"]?.if).toBe(
      "needs.ci-plan.outputs.reused_run == '' && needs.ci-plan.outputs.code_required == 'false'",
    );
    expect(jobs["warm-main-caches"]?.needs).toBe(PLAN_JOB);
    expect(jobs["warm-main-caches"]?.if).toBe(
      "github.event_name == 'push' && needs.ci-plan.outputs.reused_run != '' && needs.ci-plan.outputs.warm_caches == 'true'",
    );
  });

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
      expect([id, job.needs]).toEqual([
        id,
        id === "consumer-scenarios-minimum" ? [PLAN_JOB, "consumer-scenarios"] : PLAN_JOB,
      ]);
      const condition = job.if;
      if (PR_SCOPED_JOBS.has(id)) {
        expect(
          typeof condition === "string" && condition.startsWith(PROPERTY_AREAS_CONDITION),
        ).toBe(true);
      }
      let areaGate = condition;
      if (typeof condition === "string") {
        if (PR_SCOPED_JOBS.has(id)) areaGate = condition.slice(PROPERTY_AREAS_CONDITION.length);
        else if (id === "docs-format")
          areaGate = condition.slice("needs.ci-plan.outputs.reused_run == '' && ".length);
      }
      const gate = typeof areaGate === "string" ? GATE.exec(areaGate) : null;
      expect([id, gate !== null]).toEqual([id, true]);
      const output = gate?.[1] ?? "";
      expect([id, output in planOutputs]).toEqual([id, true]);
      if (gate?.[2] === "true") gated.add(output);
    }
    // No area is planned for nothing.
    expect([...gated].toSorted()).toEqual(
      [...areas.map((area) => `${area}_required`), "api_required"].toSorted(),
    );
  });
});
