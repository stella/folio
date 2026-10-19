import { scanTestTimeout } from "../test/scan-timeout";

import { setDefaultTimeout, describe, expect, test } from "bun:test";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

setDefaultTimeout(scanTestTimeout(30_000));

const repo = path.resolve(import.meta.dir, "..");
const workflow = Bun.YAML.parse(readFileSync(path.join(repo, ".github/workflows/ci.yml"), "utf8"));
const proofs = Bun.YAML.parse(
  readFileSync(path.join(repo, ".github/workflows/typecheck-proofs.yml"), "utf8"),
);
const detector = workflow.jobs["ci-plan"].steps.find((step) => step.id === "toolchain");
const runDetector = ({ result, base = "event-base" }) => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-toolchain-plan-"));
  const output = path.join(directory, "output");
  writeFileSync(output, "");
  try {
    const command = `bunx() { [[ "$1" == "--no-install" && "$2" == "stll-toolchain-changed" && "$3" == "--since" && "$4" == "$TOOLCHAIN_BASE" ]] || return 2; printf '%s\\n' "$DETECTOR_RESULT"; }; ${detector.run}`;
    const execution = Bun.spawnSync(["bash", "-c", command], {
      cwd: repo,
      env: {
        ...process.env,
        TOOLCHAIN_BASE: base,
        GITHUB_OUTPUT: output,
        DETECTOR_RESULT: JSON.stringify(result),
      },
    });
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

  test("compiler proofs run only through the manual workflow", () => {
    expect(proofs.on).toEqual({ workflow_dispatch: null });
    expect(proofs.permissions).toEqual({ contents: "read" });
    expect(proofs.env.TOOLCHAIN_BASE).toBe("unavailable-toolchain-base");
    const proofCommands = proofs.jobs["typecheck-proofs"].steps
      .filter((step) => /stll-typecheck-probe|typecheck:compiler-parity/.test(step.run ?? ""))
      .map((step) => step.run);
    expect(proofCommands).toEqual([
      "bunx --no-install stll-typecheck-probe --project packages/core/tsconfig.build.json --seed-dir packages/core/src -- bun check -p packages/core/tsconfig.build.json",
      'bun run typecheck:compiler-parity --changed-since "$TOOLCHAIN_BASE"',
    ]);
    for (const file of readdirSync(path.join(repo, ".github/workflows"))) {
      if (!/\.ya?ml$/.test(file) || file === "typecheck-proofs.yml") continue;
      const automatic = Bun.YAML.parse(
        readFileSync(path.join(repo, ".github/workflows", file), "utf8"),
      );
      for (const job of Object.values(automatic.jobs ?? {})) {
        for (const step of job.steps ?? []) {
          expect(step.run ?? "").not.toMatch(/stll-typecheck-probe|typecheck:compiler-parity/);
        }
        expect(job.uses ?? "").not.toContain("typecheck-proofs.yml");
      }
    }
    expect(workflow.jobs.typecheck.steps.filter((step) => step.run)).toEqual([
      { run: "bun install --frozen-lockfile" },
      { name: "Typecheck", run: "bun run typecheck" },
    ]);
  });

  test("compiler-work counters run only as a manual proof after parity", () => {
    expect(workflow.jobs["typecheck-budget"]).toBeUndefined();
    expect(workflow.jobs["ci-plan"].outputs.typecheck_budget_required).toBeUndefined();
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
          expect(step.run ?? "").not.toMatch(/typecheck-budget\.ts|typecheck:budget/);
        }
      }
    }
  });

  test("ordinary typecheck commands use Bun and preserve Vue SFC checkers", () => {
    const rootManifest = JSON.parse(readFileSync(path.join(repo, "package.json"), "utf8"));
    const projects = Bun.spawnSync(["git", "ls-files", "**/package.json"], { cwd: repo });
    expect(projects.exitCode).toBe(0);
    const vueProjects = new Set([
      "packages/vue/package.json",
      "packages/playground-vue/package.json",
    ]);
    const retained = new Set();
    for (const file of projects.stdout.toString().trim().split("\n")) {
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
    for (const [, name] of rootManifest.scripts.typecheck.matchAll(/bun run (\S+)/g)) {
      expect(rootManifest.scripts[name]).toMatch(/^bun check --no-pretty --all --project=\S+$/);
    }
  });

  test("parity selects exactly the Bun projects declared by repository scripts", () => {
    const tracked = Bun.spawnSync(["git", "ls-files", "package.json", "**/package.json"], {
      cwd: repo,
    });
    expect(tracked.exitCode).toBe(0);
    const projects = new Set();
    for (const file of tracked.stdout.toString().trim().split("\n")) {
      const manifest = JSON.parse(readFileSync(path.join(repo, file), "utf8"));
      for (const command of Object.values(manifest.scripts ?? {})) {
        if (typeof command !== "string") continue;
        const checked = /^bun check --no-pretty --all --project=(\S+)$/.exec(command);
        if (!checked) continue;
        const project = checked[1];
        projects.add(
          path
            .relative(repo, path.resolve(repo, path.dirname(file), project))
            .replaceAll(path.sep, "/"),
        );
      }
    }
    const manifest = JSON.parse(readFileSync(path.join(repo, "package.json"), "utf8"));
    const selected = [
      ...manifest.scripts["typecheck:compiler-parity"].matchAll(/--project (\S+)/g),
    ].map((match) => match[1]);
    expect(selected.toSorted()).toEqual([...projects].toSorted());
    expect(selected.length).toBe(new Set(selected).size);
    const probe = proofs.jobs["typecheck-proofs"].steps.find((step) =>
      step.run?.includes("stll-typecheck-probe"),
    );
    expect(projects.has(/--project (\S+)/.exec(probe.run)?.[1])).toBe(true);
  });

  test("every Bun project install root is installed before compiler comparison", () => {
    const manifest = JSON.parse(readFileSync(path.join(repo, "package.json"), "utf8"));
    const roots = new Set();
    for (const [, project] of manifest.scripts["typecheck:compiler-parity"].matchAll(
      /--project (\S+)/g,
    )) {
      let directory = path.dirname(path.join(repo, project));
      while (!existsSync(path.join(directory, "bun.lock"))) {
        if (directory === repo) throw new Error(`No install root for ${project}`);
        directory = path.dirname(directory);
      }
      roots.add(directory);
    }
    const steps = proofs.jobs["typecheck-proofs"].steps;
    const comparison = steps.findIndex((step) => step.run?.includes("typecheck:compiler-parity"));
    expect(comparison).toBeGreaterThan(0);
    const installed = new Set(
      steps
        .slice(0, comparison)
        .filter((step) => step.run === "bun install --frozen-lockfile")
        .map((step) => path.resolve(repo, step["working-directory"] ?? ".")),
    );
    expect([...installed].toSorted()).toEqual([...roots].toSorted());
  });

  test("a committed Bun change selects full suites against the event base", () => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-toolchain-history-"));
    const output = path.join(directory, "output");
    const git = (...args) => {
      const result = Bun.spawnSync(["git", ...args], { cwd: directory });
      if (result.exitCode !== 0) throw new Error(result.stderr.toString());
      return result.stdout.toString().trim();
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
        const execution = Bun.spawnSync(
          ["bash", "-c", `bunx() { shift 2; "$SHARED_DETECTOR" "$@"; }; ${detector.run}`],
          {
            cwd: directory,
            env: {
              ...process.env,
              TOOLCHAIN_BASE: since,
              GITHUB_OUTPUT: output,
              SHARED_DETECTOR: path.join(repo, "node_modules/.bin/stll-toolchain-changed"),
            },
          },
        );
        expect({
          status: execution.exitCode,
          output: readFileSync(output, "utf8"),
          stderr: execution.stderr.toString(),
        }).toEqual({ status: 0, output: `changed=${changed}\n`, stderr: "" });
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test("unchanged is the only detector result that retains scoped fast suites", () => {
    expect(runDetector({ result: { status: "compared", changed: false, tools: [] } })).toEqual({
      status: 0,
      output: "changed=false\n",
    });
    for (const result of [
      { status: "compared", changed: true, tools: ["bun"] },
      { status: "unreadable", changed: true, error: "unavailable base" },
    ]) {
      expect(runDetector({ result })).toEqual({ status: 0, output: "changed=true\n" });
    }
    expect(runDetector({ result: { status: "compared", changed: "false" } }).status).not.toBe(0);
  });

  test("an unreadable detector result is reported as an error annotation", () => {
    const command = `bunx() { printf '%s\\n' "$DETECTOR_RESULT"; }; ${detector.run}`;
    const directory = mkdtempSync(path.join(tmpdir(), "folio-toolchain-plan-"));
    try {
      const run = (result) =>
        Bun.spawnSync(["bash", "-c", command], {
          cwd: repo,
          env: {
            ...process.env,
            TOOLCHAIN_BASE: "event-base",
            GITHUB_OUTPUT: path.join(directory, "output"),
            DETECTOR_RESULT: JSON.stringify(result),
          },
        }).stdout.toString();
      expect(run({ status: "unreadable", changed: true, error: "unavailable base" })).toContain(
        "::error title=Toolchain change detection unreadable::unavailable base; every job runs.",
      );
      expect(run({ status: "compared", changed: true, tools: ["bun"] })).not.toContain("::error");
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  // The shared detector fails open on anything it cannot read, which silently runs
  // every job; this repository's own declarations must stay readable.
  test("the shared detector reads this repository's toolchain declarations", () => {
    const run = Bun.spawnSync(
      [
        process.execPath,
        path.join(repo, "node_modules/.bin/stll-toolchain-changed"),
        "--since",
        "HEAD",
      ],
      { cwd: repo },
    );
    expect(run.exitCode).toBe(0);
    const result = JSON.parse(run.stdout.toString());
    expect({ status: result.status, error: result.error }).toEqual({
      status: "compared",
      error: undefined,
    });
  });
});

describe("required shared toolchain policy", () => {
  const lint = workflow.jobs.lint;
  const evaluation = workflow.jobs["ci-result"].steps.find(
    (step) => step.name === "Evaluate CI outcome",
  );

  test("the existing lint job runs the policy without error tolerance unless CI is reused", () => {
    expect(lint.needs).toBe("ci-plan");
    expect(lint.if).toBe("needs.ci-plan.outputs.reused_run == ''");
    expect(lint["continue-on-error"]).toBeUndefined();
    const guards = lint.steps.filter((step) => step.run === "bun run check:toolchain");
    expect(guards).toEqual([{ name: "Shared toolchain policy", run: "bun run check:toolchain" }]);
    expect(JSON.parse(evaluation.env.JOB_SCOPES).lint).toEqual({ reuse: "forbidden" });
  });

  test("the required outcome rejects unsuccessful lint even for scoped PRs", () => {
    const scopes = JSON.parse(evaluation.env.JOB_SCOPES);
    for (const codeRequired of ["true", "false"]) {
      for (const lintResult of ["success", "failure", "skipped", "cancelled"]) {
        const needs = Object.fromEntries(
          Object.keys(scopes).map((name) => [name, { result: "success" }]),
        );
        needs.lint.result = lintResult;
        const execution = Bun.spawnSync(["bash", "-c", evaluation.run], {
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
        });
        expect(execution.exitCode).toBe(lintResult === "success" ? 0 : 1);
        if (lintResult !== "success") expect(execution.stdout.toString()).toContain("lint (");
      }
    }
  });

  test("the installed policy rejects an unpinned action in a tracked workflow", () => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-toolchain-policy-"));
    const policy = JSON.parse(
      readFileSync(path.join(repo, "node_modules/@stll/oxlint-config/toolchain.json"), "utf8"),
    );
    const action = policy.actions["actions/checkout"];
    const workflowFile = path.join(directory, ".github/workflows/example.yml");
    const checker = path.join(repo, "node_modules/.bin/stll-toolchain-check");
    const check = () => Bun.spawnSync(["bun", checker], { cwd: directory });
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
      writeWorkflow(action.sha);
      for (const args of [
        ["init", "--quiet"],
        ["add", ".github"],
      ]) {
        expect(Bun.spawnSync(["git", ...args], { cwd: directory }).exitCode).toBe(0);
      }
      const approved = check();
      expect({ status: approved.exitCode, diagnostics: approved.stderr.toString() }).toEqual({
        status: 0,
        diagnostics: "",
      });
      writeWorkflow(action.version);
      const unpinned = check();
      expect(unpinned.exitCode).toBe(1);
      expect(unpinned.stderr.toString()).toContain(
        "[action-pins] remote actions must use a full commit SHA",
      );
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});

describe("published Node minimum", () => {
  const configuration = JSON.parse(readFileSync(path.join(repo, "stll-toolchain.json"), "utf8"));
  const floor = configuration.engineFloors.at(0);

  test("the Hawk exception matches the workflow and invoked toolchain", () => {
    const exception = configuration.rustExceptions.at(0);
    expect(exception).toEqual({
      pin: "1.97.1",
      reason: "cargo-hawk requires Rust 1.97.1; remove when Hawk supports the shared Rust pin.",
      tracking: "https://github.com/stella/tooling/issues/99",
      expires: "2026-11-11",
    });
    expect(workflow.env.HAWK_RUST_TOOLCHAIN).toBe(exception.pin);
    const manifest = JSON.parse(readFileSync(path.join(repo, "package.json"), "utf8"));
    expect(manifest.scripts["rust:hawk"]).toStartWith(`cargo +${exception.pin} hawk check `);
    expect(manifest.scripts["kernel:check"]).toContain("bun run rust:hawk");
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
    expect(current.steps.find((step) => step.uses?.startsWith("actions/setup-node@")).with).toEqual(
      { "node-version-file": ".node-version" },
    );
    expect(current.steps.filter((step) => step.run?.includes("--pack-only"))).toEqual([
      {
        name: "Build consumer tarballs with the shared runtime",
        run: "bun scripts/consumer-scenarios.ts --pack-only test-results/consumer-tarballs",
      },
    ]);
    const upload = current.steps.find((step) => step.with?.name === "consumer-scenarios-tarballs");
    const download = minimum.steps.find((step) => step.with?.name === upload.with.name);
    expect(download.with.path).toBe(upload.with.path);
    expect(
      minimum.steps
        .filter((step) => step.run?.includes("consumer-scenarios.ts"))
        .map((step) => step.run),
    ).toEqual([`bun scripts/consumer-scenarios.ts --typecheck --tarballs ${download.with.path}`]);
    expect(JSON.parse(evaluationForMinimum().env.JOB_SCOPES)[floor.job]).toEqual({
      area: "packaged_consumer_required",
    });
    const root = JSON.parse(readFileSync(path.join(repo, "package.json"), "utf8"));
    expect(root.scripts["check:toolchain"]).toBe(
      "stll-toolchain-check && bun scripts/consumer-node-minimum.mjs --check",
    );
  });

  const evaluationForMinimum = () =>
    workflow.jobs["ci-result"].steps.find((step) => step.name === "Evaluate CI outcome");

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
      const execution = Bun.spawnSync(["bash", "-c", evaluation.run], {
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
      });
      expect(execution.exitCode).toBe(result === "success" ? 0 : 1);
      if (result !== "success") expect(execution.stdout.toString()).toContain(`${floor.job} (`);
    }
  });

  test.each(
    Array.from({ length: 4 }, (_, majorOffset) => 22 + majorOffset).flatMap((major) =>
      [
        [0, 0],
        [1, 2],
        [18, 4],
      ].map(([minor, patch]) => [major, minor, patch]),
    ),
  )("engine minimum %i.%i.%i is generated exactly and idempotently", (major, minor, patch) => {
    const directory = mkdtempSync(path.join(tmpdir(), "folio-node-minimum-"));
    const checker = path.join(repo, "node_modules/.bin/stll-toolchain-check");
    const projection = path.join(repo, "scripts/consumer-node-minimum.mjs");
    const shared = () => Bun.spawnSync([process.execPath, checker], { cwd: directory });
    const project = (mode) =>
      Bun.spawnSync([process.execPath, projection, mode], { cwd: directory });
    const check = () => {
      const result = shared();
      return result.exitCode === 0 ? project("--check") : result;
    };
    const manifest = path.join(directory, floor.package, "package.json");
    const workflowFile = path.join(directory, floor.workflow);
    const setRange = (range) =>
      writeFileSync(manifest, JSON.stringify({ name: "@example/cli", engines: { node: range } }));
    try {
      mkdirSync(path.dirname(manifest), { recursive: true });
      mkdirSync(path.dirname(workflowFile), { recursive: true });
      for (const file of [
        ".github/dependabot.yml",
        ".node-version",
        "stll-toolchain.json",
        ".agents/cloud-setup.sh",
        ".claude/settings.json",
        "rust-toolchain.toml",
        "Cargo.lock",
        floor.workflow,
      ]) {
        mkdirSync(path.dirname(path.join(directory, file)), { recursive: true });
        writeFileSync(path.join(directory, file), readFileSync(path.join(repo, file), "utf8"));
      }
      writeFileSync(
        path.join(directory, "package.json"),
        JSON.stringify({
          private: true,
          packageManager:
            "bun@" +
            JSON.parse(
              readFileSync(
                path.join(repo, "node_modules/@stll/oxlint-config/toolchain.json"),
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
        expect(Bun.spawnSync(["git", ...args], { cwd: directory }).exitCode).toBe(0);
      const initial = check();
      expect({ status: initial.exitCode, diagnostic: initial.stderr.toString() }).toEqual({
        status: 0,
        diagnostic: "",
      });
      const original = readFileSync(workflowFile, "utf8");
      const selected = Bun.YAML.parse(original).jobs[floor.job].steps.find((step) =>
        step.uses?.startsWith("actions/setup-node@"),
      ).with["node-version"];
      const drifted = original.replace(`node-version: "${selected}"`, 'node-version: "22.23.3"');
      expect(drifted).not.toBe(original);
      writeFileSync(workflowFile, drifted);
      expect(shared().exitCode).toBe(0);
      const drift = check();
      expect(drift.exitCode).toBe(1);
      expect(drift.stderr.toString()).toContain("[consumer-node-minimum]");
      const version = `${major}.${minor}.${patch}`;
      setRange(`>=${version}`);
      expect(check().exitCode).toBe(1);
      const generated = project("--write");
      expect({ status: generated.exitCode, diagnostic: generated.stderr.toString() }).toEqual({
        status: 0,
        diagnostic: "",
      });
      expect(
        Bun.YAML.parse(readFileSync(workflowFile, "utf8")).jobs[floor.job].steps.find((step) =>
          step.uses?.startsWith("actions/setup-node@"),
        ).with["node-version"],
      ).toBe(version);
      expect(check().exitCode).toBe(0);
      const once = readFileSync(workflowFile, "utf8");
      expect(project("--write").exitCode).toBe(0);
      expect(readFileSync(workflowFile, "utf8")).toBe(once);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});
