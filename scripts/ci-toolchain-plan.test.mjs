import { describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const repo = path.resolve(import.meta.dir, "..");
const workflow = Bun.YAML.parse(readFileSync(path.join(repo, ".github/workflows/ci.yml"), "utf8"));
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
    expect(
      workflow.jobs.typecheck.steps.some(
        (step) =>
          step.run === 'bun run typecheck:compiler-parity --changed-since "$TOOLCHAIN_BASE"',
      ),
    ).toBe(true);
  });

  test("parity selects exactly the native TS7 projects declared by repository scripts", () => {
    const tracked = Bun.spawnSync(["git", "ls-files", "package.json", "**/package.json"], {
      cwd: repo,
    });
    expect(tracked.exitCode).toBe(0);
    const projects = new Set();
    for (const file of tracked.stdout.toString().trim().split("\n")) {
      const manifest = JSON.parse(readFileSync(path.join(repo, file), "utf8"));
      for (const command of Object.values(manifest.scripts ?? {})) {
        if (typeof command !== "string") continue;
        const native = /^bun (?:\.\.\/)*scripts\/tsc-native\.ts(?: (.*))?$/.exec(command);
        if (!native) continue;
        const args = (native[1] ?? "").split(/\s+/).filter(Boolean);
        let project = "tsconfig.json";
        for (let index = 0; index < args.length; index += 1) {
          if (args[index] === "--noEmit") continue;
          if (args[index] !== "-p" || !args[index + 1])
            throw new Error(`Unsupported native compiler arguments in ${file}`);
          project = args[index + 1];
          index += 1;
        }
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
    const probe = workflow.jobs.typecheck.steps.find((step) =>
      step.run?.includes("stll-typecheck-probe"),
    );
    expect(projects.has(/--project (\S+)/.exec(probe.run)?.[1])).toBe(true);
  });

  test("every native project install root is installed before compiler comparison", () => {
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
    const steps = workflow.jobs.typecheck.steps;
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
});
