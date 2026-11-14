import { describe, expect, setDefaultTimeout, test } from "bun:test";
import {
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

// The largest test here runs 13 children one after another.
setDefaultTimeout(subprocessTestTimeout(13));

const repo = path.resolve(import.meta.dir, "..");

/**
 * Runs a command in a fixture repository as it would run outside CI.
 *
 * The outer run's CI context (its event, the event's base commit, whether it
 * is a push or a schedule) describes this repository, not the fixture. A
 * checker that inherits it judges the fixture's dated entries against a base
 * commit the fixture does not have, so the fixture's verdict would change with
 * the calendar and the outer event instead of with what the test wrote.
 */
const fixtureEnvironment = Object.fromEntries(
  Object.entries(process.env).filter(
    ([name]) => name !== "CI" && !name.startsWith("GITHUB_"),
  ),
);
const runInFixture = (command, directory) =>
  expectExited(runChild(command, { cwd: directory, env: fixtureEnvironment }));

const workflow = Bun.YAML.parse(
  readFileSync(path.join(repo, ".github/workflows/ci.yml"), "utf8"),
);
const proofs = Bun.YAML.parse(
  readFileSync(
    path.join(repo, ".github/workflows/typecheck-proofs.yml"),
    "utf8",
  ),
);
const detector = workflow.jobs["ci-plan"].steps.find(
  (step) => step.id === "toolchain",
);
const runDetector = ({ result, base = "event-base" }) => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-toolchain-plan-"));
  const output = path.join(directory, "output");
  writeFileSync(output, "");
  try {
    const command = `bunx() { [[ "$1" == "--no-install" && "$2" == "stll-toolchain-changed" && "$3" == "--since" && "$4" == "$TOOLCHAIN_BASE" ]] || return 2; printf '%s\\n' "$DETECTOR_RESULT"; }; ${detector.run}`;
    const execution = expectExited(
      runChild(["bash", "-c", command], {
        cwd: repo,
        env: {
          ...process.env,
          TOOLCHAIN_BASE: base,
          GITHUB_OUTPUT: output,
          DETECTOR_RESULT: JSON.stringify(result),
        },
      }),
    );
    return { status: execution.exitCode, output: readFileSync(output, "utf8") };
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

describe("shared toolchain CI selection", () => {
  test("every event compares its real base and every selected checkout keeps history", () => {
    expect(workflow.env.TOOLCHAIN_BASE).toBe(
      "${{ (github.event_name == 'pull_request' && github.event.pull_request.base.sha) || (github.event_name == 'merge_group' && github.event.merge_group.base_sha) || (github.event_name == 'push' && github.event.before) || 'unavailable-toolchain-base' }}",
    );
    for (const name of ["ci-plan", "typecheck"]) {
      const checkout = workflow.jobs[name].steps.find((step) =>
        step.uses?.startsWith("actions/checkout@"),
      );
      expect(checkout.with["fetch-depth"]).toBe(0);
    }
  });

  // A hand-started run names no base; the detector reads the one the shared
  // comparison-base owner resolves instead of reporting an unreadable base.
  test("the detector compares with the base the owner resolves", () => {
    const steps = workflow.jobs["ci-plan"].steps;
    const resolver = steps.findIndex((step) => step.id === "toolchain-base");
    expect(steps[resolver].run).toBe(
      'bun scripts/lib/comparison-base.ts >> "$GITHUB_OUTPUT"',
    );
    expect(steps[resolver].env).toBeUndefined();
    expect(steps.indexOf(detector)).toBe(resolver + 1);
    expect(detector.env).toEqual({
      TOOLCHAIN_BASE: "${{ steps.toolchain-base.outputs.base }}",
    });
  });

  test("compiler proofs run only through the manual workflow", () => {
    expect(proofs.on).toEqual({ workflow_dispatch: null });
    expect(proofs.permissions).toEqual({ contents: "read" });
    const proofCommands = proofs.jobs["typecheck-proofs"].steps
      .filter((step) => /stll-typecheck-probe/.test(step.run ?? ""))
      .map((step) => step.run);
    expect(proofCommands).toEqual([
      "bunx --no-install stll-typecheck-probe --project packages/core/tsconfig.build.json --seed-dir packages/core/src -- bun check -p packages/core/tsconfig.build.json",
    ]);
    for (const file of readdirSync(path.join(repo, ".github/workflows"))) {
      if (!/\.ya?ml$/.test(file) || file === "typecheck-proofs.yml") continue;
      const automatic = Bun.YAML.parse(
        readFileSync(path.join(repo, ".github/workflows", file), "utf8"),
      );
      for (const job of Object.values(automatic.jobs ?? {})) {
        for (const step of job.steps ?? []) {
          expect(step.run ?? "").not.toMatch(/stll-typecheck-probe/);
        }
        expect(job.uses ?? "").not.toContain("typecheck-proofs.yml");
      }
    }
    expect(workflow.jobs.typecheck.steps.filter((step) => step.run)).toEqual([
      { run: "bun install --frozen-lockfile" },
      {
        name: "Typecheck",
        run: 'bun scripts/ci-plan.ts --base "$TYPECHECK_BASE" --head "$GITHUB_SHA"',
        env: {
          TYPECHECK_BASE:
            "${{ (github.event_name == 'workflow_dispatch' && '0000000000000000000000000000000000000000') || github.event.pull_request.base.sha || github.event.merge_group.base_sha || github.event.before }}",
          TYPECHECK_PUSH: "${{ github.event_name == 'push' }}",
        },
      },
    ]);
  });

  test("compiler-work counters run only as a manual proof after the probe", () => {
    expect(workflow.jobs["typecheck-budget"]).toBeUndefined();
    expect(
      workflow.jobs["ci-plan"].outputs.typecheck_budget_required,
    ).toBeUndefined();
    const budgetJob = proofs.jobs["typecheck-budget"];
    expect(budgetJob.needs).toBe("typecheck-proofs");
    expect(budgetJob.if).toBe("always()");
    expect(budgetJob.steps.filter((step) => step.run)).toEqual([
      { run: "bun scripts/ci-scoped-install.ts typecheck-budget" },
      {
        name: "Guard deterministic compiler workload",
        run: "bun scripts/typecheck-budget.ts --self-test\nbun run typecheck:budget\n",
      },
    ]);
    for (const file of readdirSync(path.join(repo, ".github/workflows"))) {
      if (!/\.ya?ml$/.test(file) || file === "typecheck-proofs.yml") continue;
      const automatic = Bun.YAML.parse(
        readFileSync(path.join(repo, ".github/workflows", file), "utf8"),
      );
      for (const job of Object.values(automatic.jobs ?? {})) {
        for (const step of job.steps ?? []) {
          expect(step.run ?? "").not.toMatch(
            /typecheck-budget\.ts|typecheck:budget/,
          );
        }
      }
    }
  });

  test("ordinary typecheck commands use Bun and preserve Vue SFC checkers", () => {
    const rootManifest = JSON.parse(
      readFileSync(path.join(repo, "package.json"), "utf8"),
    );
    const projects = expectExited(
      runChild(["git", "ls-files", "**/package.json"], { cwd: repo }),
    );
    expect(projects.exitCode).toBe(0);
    const vueProjects = new Set([
      "packages/vue/package.json",
      "packages/playground-vue/package.json",
    ]);
    const retained = new Set();
    for (const file of projects.stdout.trim().split("\n")) {
      const manifest = JSON.parse(readFileSync(path.join(repo, file), "utf8"));
      const command = manifest.scripts?.typecheck;
      if (command === undefined) continue;
      if (vueProjects.has(file)) {
        expect(command).toMatch(/^vue-tsc --noEmit(?: -p \S+)?$/);
        retained.add(file);
        continue;
      }
      expect(command).toMatch(/^bun check --no-pretty --all --project=\S+$/);
    }
    expect([...retained].toSorted()).toEqual([...vueProjects].toSorted());
    for (const [, name] of rootManifest.scripts.typecheck.matchAll(
      /bun run (\S+)/g,
    )) {
      expect(rootManifest.scripts[name]).toMatch(
        /^bun check --no-pretty --all --project=\S+$/,
      );
    }
  });

  test("a committed Bun change selects full suites against the event base", () => {
    const directory = mkdtempSync(
      path.join(tmpdir(), "folio-toolchain-history-"),
    );
    const output = path.join(directory, "output");
    const git = (...args) => {
      const result = expectExited(
        runChild(["git", ...args], { cwd: directory }),
      );
      if (result.exitCode !== 0) throw new Error(result.stderr);
      return result.stdout.trim();
    };
    const commit = (version) => {
      writeFileSync(
        path.join(directory, "package.json"),
        JSON.stringify({ private: true, packageManager: `bun@${version}` }),
      );
      git("add", "package.json");
      git(
        "-c",
        "user.name=Fixture",
        "-c",
        "user.email=fixture@example.test",
        "-c",
        "commit.gpgsign=false",
        "commit",
        "-m",
        "chore: set fixture runtime",
      );
      return git("rev-parse", "HEAD");
    };
    try {
      git("init", "--quiet");
      const base = commit("1.4.2");
      const head = commit("1.4.3");
      for (const [since, changed] of [
        [base, "true"],
        [head, "false"],
        ["unavailable-toolchain-base", "true"],
      ]) {
        writeFileSync(output, "");
        const execution = expectExited(
          runChild(
            [
              "bash",
              "-c",
              `bunx() { shift 2; "$SHARED_DETECTOR" "$@"; }; ${detector.run}`,
            ],
            {
              cwd: directory,
              env: {
                ...process.env,
                TOOLCHAIN_BASE: since,
                GITHUB_OUTPUT: output,
                SHARED_DETECTOR: path.join(
                  repo,
                  "node_modules/.bin/stll-toolchain-changed",
                ),
              },
            },
          ),
        );
        expect({
          status: execution.exitCode,
          output: readFileSync(output, "utf8"),
          stderr: execution.stderr,
        }).toEqual({ status: 0, output: `changed=${changed}\n`, stderr: "" });
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("unchanged is the only detector result that retains scoped fast suites", () => {
    expect(
      runDetector({
        result: { status: "compared", changed: false, tools: [] },
      }),
    ).toEqual({
      status: 0,
      output: "changed=false\n",
    });
    for (const result of [
      { status: "compared", changed: true, tools: ["bun"] },
      { status: "unreadable", changed: true, error: "unavailable base" },
    ]) {
      expect(runDetector({ result })).toEqual({
        status: 0,
        output: "changed=true\n",
      });
    }
    expect(
      runDetector({ result: { status: "compared", changed: "false" } }).status,
    ).not.toBe(0);
  });

  test("an unreadable detector result is reported as an error annotation", () => {
    const command = `bunx() { printf '%s\\n' "$DETECTOR_RESULT"; }; ${detector.run}`;
    const directory = mkdtempSync(path.join(tmpdir(), "folio-toolchain-plan-"));
    try {
      const run = (result) =>
        expectExited(
          runChild(["bash", "-c", command], {
            cwd: repo,
            env: {
              ...process.env,
              TOOLCHAIN_BASE: "event-base",
              GITHUB_OUTPUT: path.join(directory, "output"),
              DETECTOR_RESULT: JSON.stringify(result),
            },
          }),
        ).stdout;
      expect(
        run({ status: "unreadable", changed: true, error: "unavailable base" }),
      ).toContain(
        "::error title=Toolchain change detection unreadable::unavailable base; every job runs.",
      );
      expect(
        run({ status: "compared", changed: true, tools: ["bun"] }),
      ).not.toContain("::error");
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  // The shared detector fails open on anything it cannot read, which silently runs
  // every job; this repository's own declarations must stay readable.
  test("the shared detector reads this repository's toolchain declarations", () => {
    const run = expectExited(
      runChild(
        [
          process.execPath,
          path.join(repo, "node_modules/.bin/stll-toolchain-changed"),
          "--since",
          "HEAD",
        ],
        { cwd: repo },
      ),
    );
    expect(run.exitCode).toBe(0);
    const result = JSON.parse(run.stdout);
    expect({ status: result.status, error: result.error }).toEqual({
      status: "compared",
      error: undefined,
    });
  });
});

describe("required shared toolchain policy", () => {
  // Its dated exceptions expire on a date, so it runs beside lint.
  const policy = workflow.jobs["dated-policies"];
  const evaluation = workflow.jobs["ci-result"].steps.find(
    (step) => step.name === "Evaluate CI outcome",
  );

  test("the dated policy job runs the policy without error tolerance unless CI is reused", () => {
    expect(policy.needs).toBe("ci-plan");
    expect(policy.if).toBe("needs.ci-plan.outputs.reused_run == ''");
    expect(policy["continue-on-error"]).toBeUndefined();
    const guards = policy.steps.filter(
      (step) => step.run === "bun run check:toolchain",
    );
    expect(guards).toEqual([
      { name: "Shared toolchain policy", run: "bun run check:toolchain" },
    ]);
    expect(JSON.parse(evaluation.env.JOB_SCOPES)["dated-policies"]).toEqual({
      reuse: "forbidden",
    });
  });

  test("the required outcome rejects unsuccessful lint even for scoped PRs", () => {
    const scopes = JSON.parse(evaluation.env.JOB_SCOPES);
    for (const codeRequired of ["true", "false"]) {
      for (const lintResult of ["success", "failure", "skipped", "cancelled"]) {
        const needs = Object.fromEntries(
          Object.keys(scopes).map((name) => [name, { result: "success" }]),
        );
        needs.lint.result = lintResult;
        const execution = expectExited(
          runChild(["bash", "-c", evaluation.run], {
            cwd: repo,
            env: {
              ...process.env,
              EVENT: "pull_request",
              PLAN_RESULT: "success",
              NEEDS: JSON.stringify(needs),
              PLAN: JSON.stringify({
                trusted: "true",
                suite_depth: "fast",
                toolchain_changed: "false",
                code_required: codeRequired,
              }),
              JOB_SCOPES: evaluation.env.JOB_SCOPES,
            },
          }),
        );
        expect(execution.exitCode).toBe(lintResult === "success" ? 0 : 1);
        if (lintResult !== "success")
          expect(execution.stdout).toContain("lint (");
      }
    }
  });

  test("the installed policy rejects an unpinned action in a tracked workflow", () => {
    const directory = mkdtempSync(
      path.join(tmpdir(), "folio-toolchain-policy-"),
    );
    const policy = JSON.parse(
      readFileSync(
        path.join(repo, "node_modules/@stll/oxlint-config/toolchain.json"),
        "utf8",
      ),
    );
    const action = policy.actions["actions/checkout"];
    const workflowFile = path.join(directory, ".github/workflows/example.yml");
    const checker = path.join(repo, "node_modules/.bin/stll-toolchain-check");
    const check = () => runInFixture(["bun", checker], directory);
    const writeWorkflow = (selector) =>
      writeFileSync(
        workflowFile,
        `name: Example\non: push\njobs:\n  example:\n    runs-on: ubuntu-24.04\n    steps:\n      - uses: actions/checkout@${selector} # ${action.version}\n`,
      );
    try {
      mkdirSync(path.dirname(workflowFile), { recursive: true });
      writeFileSync(
        path.join(directory, ".github/dependabot.yml"),
        readFileSync(path.join(repo, ".github/dependabot.yml"), "utf8"),
      );
      // The repository's own Dependabot ignores, which the copied
      // configuration carries; nothing else of the manifest applies here.
      const { dependabotIgnores } = JSON.parse(
        readFileSync(path.join(repo, "stll-toolchain.json"), "utf8"),
      );
      writeFileSync(
        path.join(directory, "stll-toolchain.json"),
        `${JSON.stringify({ dependabotIgnores }, null, 2)}\n`,
      );
      writeWorkflow(action.sha);
      for (const args of [
        ["init", "--quiet"],
        ["add", ".github", "stll-toolchain.json"],
      ]) {
        expect(
          expectExited(runChild(["git", ...args], { cwd: directory })).exitCode,
        ).toBe(0);
      }
      const approved = check();
      expect({
        status: approved.exitCode,
        diagnostics: approved.stderr,
      }).toEqual({
        status: 0,
        diagnostics: "",
      });
      writeWorkflow(action.version);
      const unpinned = check();
      expect(unpinned.exitCode).toBe(1);
      expect(unpinned.stderr).toContain(
        "[action-pins] remote actions must use a full commit SHA",
      );
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});

describe("published Node minimum", () => {
  const configuration = JSON.parse(
    readFileSync(path.join(repo, "stll-toolchain.json"), "utf8"),
  );
  const floor = configuration.engineFloors.at(0);

  // Hawk runs only with the compiler it was built with, so it builds with
  // rust-toolchain.toml and neither the script nor the workflow names another.
  test("Hawk builds and runs with the repository's Rust pin", () => {
    expect(configuration.rustExceptions).toBeUndefined();
    const manifest = JSON.parse(
      readFileSync(path.join(repo, "package.json"), "utf8"),
    );
    expect(manifest.scripts["rust:hawk"]).toStartWith("cargo hawk check ");
    expect(manifest.scripts["kernel:check"]).toContain("bun run rust:hawk");
    const steps = workflow.jobs["docx-kernel"].steps;
    const install = steps.find((step) => step.name === "Install pinned Hawk");
    expect(install.run).toContain(
      'RUSTC_BOOTSTRAP=1 cargo install cargo-hawk --version "=${HAWK_VERSION}" --locked',
    );
    expect(install.run).toContain("rustup component add rustc-dev\n");
    expect(install.run).not.toMatch(/(?:cargo|rustc) \+|--toolchain/u);
    const cacheKeys = steps
      .filter((step) => step.with?.path === "~/.cache/cargo-hawk")
      .map((step) => step.with.key);
    expect(cacheKeys).toHaveLength(2);
    for (const key of cacheKeys)
      expect(key).toContain("${{ hashFiles('rust-toolchain.toml') }}");
    expect(
      Object.keys(workflow.env).filter((name) => name.startsWith("HAWK_")),
    ).toEqual(["HAWK_VERSION", "HAWK_CRATE_SHA256"]);
  });

  test("the minimum consumer run owns a scoped floor and reuses default-runtime tarballs", () => {
    expect(configuration.engineFloors).toEqual([
      {
        package: "packages/cli",
        workflow: ".github/workflows/ci.yml",
        job: "consumer-scenarios-minimum",
      },
    ]);
    const minimum = workflow.jobs[floor.job];
    const current = workflow.jobs["consumer-scenarios"];
    expect(minimum.needs).toEqual(["ci-plan", "consumer-scenarios"]);
    expect(minimum.if).toBe(current.if);
    expect(
      current.steps.find((step) => step.uses?.startsWith("actions/setup-node@"))
        .with,
    ).toEqual({ "node-version-file": ".node-version" });
    expect(
      current.steps.filter((step) => step.run?.includes("--pack-only")),
    ).toEqual([
      {
        name: "Build consumer tarballs with the shared runtime",
        run: "bun scripts/consumer-scenarios.ts --pack-only test-results/consumer-tarballs",
      },
    ]);
    const upload = current.steps.find(
      (step) => step.with?.name === "consumer-scenarios-tarballs",
    );
    const download = minimum.steps.find(
      (step) => step.with?.name === upload.with.name,
    );
    expect(download.with.path).toBe(upload.with.path);
    expect(
      minimum.steps
        .filter((step) => step.run?.includes("consumer-scenarios.ts"))
        .map((step) => step.run),
    ).toEqual([
      `bun scripts/consumer-scenarios.ts --typecheck --tarballs ${download.with.path}`,
    ]);
    expect(
      JSON.parse(evaluationForMinimum().env.JOB_SCOPES)[floor.job],
    ).toEqual({
      area: "packaged_consumer_required",
    });
    const root = JSON.parse(
      readFileSync(path.join(repo, "package.json"), "utf8"),
    );
    expect(root.scripts["check:toolchain"]).toBe(
      "stll-toolchain-check && bun scripts/consumer-node-minimum.mjs --check",
    );
  });

  const evaluationForMinimum = () =>
    workflow.jobs["ci-result"].steps.find(
      (step) => step.name === "Evaluate CI outcome",
    );

  test("the required outcome rejects a skipped or failed minimum consumer run", () => {
    const evaluation = evaluationForMinimum();
    const scopes = JSON.parse(evaluation.env.JOB_SCOPES);
    const plan = Object.fromEntries(
      Object.values(scopes)
        .filter((scope) => scope.area)
        .map((scope) => [scope.area, "true"]),
    );
    for (const result of ["success", "failure", "skipped", "cancelled"]) {
      const needs = Object.fromEntries(
        Object.keys(scopes).map((job) => [job, { result: "success" }]),
      );
      needs[floor.job].result = result;
      const execution = expectExited(
        runChild(["bash", "-c", evaluation.run], {
          cwd: repo,
          env: {
            ...process.env,
            EVENT: "merge_group",
            PLAN_RESULT: "success",
            NEEDS: JSON.stringify(needs),
            PLAN: JSON.stringify({
              ...plan,
              trusted: "true",
              suite_depth: "full",
              toolchain_changed: "false",
            }),
            JOB_SCOPES: evaluation.env.JOB_SCOPES,
          },
        }),
      );
      expect(execution.exitCode).toBe(result === "success" ? 0 : 1);
      if (result !== "success")
        expect(execution.stdout).toContain(`${floor.job} (`);
    }
  });

  test.each(
    Array.from({ length: 4 }, (_, majorOffset) => 22 + majorOffset).flatMap(
      (major) =>
        [
          [0, 0],
          [1, 2],
          [18, 4],
        ].map(([minor, patch]) => [major, minor, patch]),
    ),
  )(
    "engine minimum %i.%i.%i is generated exactly and idempotently",
    (major, minor, patch) => {
      const directory = mkdtempSync(path.join(tmpdir(), "folio-node-minimum-"));
      const checker = path.join(repo, "node_modules/.bin/stll-toolchain-check");
      const projection = path.join(repo, "scripts/consumer-node-minimum.mjs");
      const shared = () => runInFixture([process.execPath, checker], directory);
      const project = (mode) =>
        runInFixture([process.execPath, projection, mode], directory);
      const check = () => {
        const result = shared();
        return result.exitCode === 0 ? project("--check") : result;
      };
      const manifest = path.join(directory, floor.package, "package.json");
      const workflowFile = path.join(directory, floor.workflow);
      const setRange = (range) =>
        writeFileSync(
          manifest,
          JSON.stringify({ name: "@example/cli", engines: { node: range } }),
        );
      try {
        mkdirSync(path.dirname(manifest), { recursive: true });
        mkdirSync(path.dirname(workflowFile), { recursive: true });
        for (const file of [
          ".github/dependabot.yml",
          ".node-version",
          "bunfig.toml",
          "stll-toolchain.json",
          ".agents/cloud-setup.sh",
          ".claude/settings.json",
          "rust-toolchain.toml",
          "Cargo.lock",
          floor.workflow,
        ]) {
          mkdirSync(path.dirname(path.join(directory, file)), {
            recursive: true,
          });
          writeFileSync(
            path.join(directory, file),
            readFileSync(path.join(repo, file), "utf8"),
          );
        }
        // The fixture root declares no TypeScript install, so the repository's
        // allowlist for its own tools does not apply to it.
        const toolchain = JSON.parse(
          readFileSync(path.join(repo, "stll-toolchain.json"), "utf8"),
        );
        delete toolchain.typescriptAllowlist;
        writeFileSync(
          path.join(directory, "stll-toolchain.json"),
          JSON.stringify(toolchain),
        );
        writeFileSync(
          path.join(directory, "package.json"),
          JSON.stringify({
            private: true,
            packageManager:
              "bun@" +
              JSON.parse(
                readFileSync(
                  path.join(
                    repo,
                    "node_modules/@stll/oxlint-config/toolchain.json",
                  ),
                  "utf8",
                ),
              ).bun,
            workspaces: ["packages/*"],
          }),
        );
        setRange(">=22");
        for (const args of [
          ["init", "--quiet"],
          ["add", "."],
        ])
          expect(
            expectExited(runChild(["git", ...args], { cwd: directory }))
              .exitCode,
          ).toBe(0);
        const initial = check();
        // Dated quarantine entries near expiry print warnings; only errors count.
        expect({
          status: initial.exitCode,
          diagnostic: initial.stderr
            .toString()
            .split("\n")
            .filter((line) => line !== "" && !line.includes("] warning: "))
            .join("\n"),
        }).toEqual({
          status: 0,
          diagnostic: "",
        });
        const original = readFileSync(workflowFile, "utf8");
        const selected = Bun.YAML.parse(original).jobs[floor.job].steps.find(
          (step) => step.uses?.startsWith("actions/setup-node@"),
        ).with["node-version"];
        const drifted = original.replace(
          `node-version: "${selected}"`,
          'node-version: "22.23.3"',
        );
        expect(drifted).not.toBe(original);
        writeFileSync(workflowFile, drifted);
        expect(shared().exitCode).toBe(0);
        const drift = check();
        expect(drift.exitCode).toBe(1);
        expect(drift.stderr).toContain("[consumer-node-minimum]");
        const version = `${major}.${minor}.${patch}`;
        setRange(`>=${version}`);
        expect(check().exitCode).toBe(1);
        const generated = project("--write");
        expect({
          status: generated.exitCode,
          diagnostic: generated.stderr,
        }).toEqual({
          status: 0,
          diagnostic: "",
        });
        expect(
          Bun.YAML.parse(readFileSync(workflowFile, "utf8")).jobs[
            floor.job
          ].steps.find((step) => step.uses?.startsWith("actions/setup-node@"))
            .with["node-version"],
        ).toBe(version);
        expect(check().exitCode).toBe(0);
        const once = readFileSync(workflowFile, "utf8");
        expect(project("--write").exitCode).toBe(0);
        expect(readFileSync(workflowFile, "utf8")).toBe(once);
      } finally {
        rmSync(directory, { recursive: true, force: true });
      }
    },
  );
});
