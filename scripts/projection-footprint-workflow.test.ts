import { afterEach, expect, test } from "bun:test";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const directories: string[] = [];
afterEach(() => {
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true });
});

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const workflowJobs = () => {
  const workflow: unknown = Bun.YAML.parse(
    readFileSync(join(import.meta.dir, "../.github/workflows/canonical-wasm.yml"), "utf8"),
  );
  if (!isRecord(workflow) || !isRecord(workflow["jobs"]))
    throw new TypeError("Canonical workflow has no jobs.");
  return workflow["jobs"];
};

const workflowStep = (name: string) => {
  for (const job of Object.values(workflowJobs())) {
    if (!isRecord(job) || !Array.isArray(job["steps"])) continue;
    for (const step of job["steps"]) {
      if (!isRecord(step) || step["name"] !== name) continue;
      if (typeof step["run"] !== "string")
        throw new TypeError(`Workflow step ${name} has no executable script.`);
      return step["run"];
    }
  }
  throw new TypeError(`Canonical workflow has no ${name} step.`);
};

// Every external tool is isolated; this exercises the actual workflow shell,
// without building Rust, installing a CLI, or accessing another repository.
const mockTool = `
import { appendFileSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
const tool = basename(process.argv[1]);
const args = process.argv.slice(2);
const root = process.env.FOOTPRINT_FIXTURE;
const state = join(root, "cli-version");
const log = (message) => appendFileSync(join(root, "calls"), message + "\\n");
const option = (name) => args.at(args.indexOf(name) + 1);
const revision = (path) => path.includes("footprint-base") ? "base" : "head";
const version = (name) => name === "base" ? process.env.BASE_VERSION : process.env.HEAD_VERSION;
const fail = (message) => { throw new Error(message); };
switch (tool) {
  case "git": {
    if (args.slice(0, 3).join(" ") !== "worktree add --detach") fail("unexpected git command");
    mkdirSync(args[3], { recursive: true });
    writeFileSync(join(args[3], "Cargo.toml"), "base manifest");
    break;
  }
  case "cargo": {
    if (args[0] === "metadata") {
      const manifest = option("--manifest-path");
      if (!manifest || !args.includes("--locked")) fail("metadata must be revision-locked");
      log("metadata:" + revision(manifest));
      console.log(JSON.stringify({ packages: [{ name: "wasm-bindgen", version: version(revision(manifest)) }] }));
    } else if (args[0] === "install") {
      log("install");
      fail("comparison must not install a CLI");
    } else if (args[0] === "rustc") {
      log("rustc:" + (args.includes("--manifest-path") ? "base" : "head"));
    } else fail("unexpected cargo command");
    break;
  }
  case "jq": {
    const metadata = JSON.parse(await Bun.stdin.text());
    for (const resolved of new Set(metadata.packages.filter((item) => item.name === "wasm-bindgen").map((item) => item.version)))
      console.log(resolved);
    break;
  }
  case "wasm-bindgen": {
    const installed = readFileSync(state, "utf8");
    if (args[0] === "--version") {
      log("bindgen:version");
      console.log("wasm-bindgen " + installed);
    }
    else {
      const name = revision(args[0]);
      if (installed !== version(name)) fail("CLI does not match " + name + " dependency");
      log("bindgen:" + name + ":" + installed);
      const output = option("--out-dir");
      mkdirSync(output, { recursive: true });
      writeFileSync(join(output, "projection_bg.wasm"), installed);
    }
    break;
  }
  case "bun": {
    if (args[0] !== "scripts/projection-footprint.ts") fail("unexpected Bun command");
    if (readFileSync(args[1], "utf8") !== process.env.BASE_VERSION ||
        readFileSync(args[2], "utf8") !== process.env.HEAD_VERSION)
      fail("comparison must receive base then head from their matching CLIs");
    log("compare:base:head");
    writeFileSync(args[3], "{}");
    break;
  }
  default: fail("unexpected mock executable");
}
`;

test.each([
  {
    name: "matching revisions share one CLI after resolving both manifests",
    head: "0.2.126",
    base: "0.2.126",
    exitCode: 0,
    calls: [
      "metadata:head",
      "metadata:base",
      "rustc:head",
      "bindgen:head:0.2.126",
      "rustc:base",
      "bindgen:base:0.2.126",
      "compare:base:head",
    ],
  },
  {
    name: "different versions fail before any build, binding, or installation",
    head: "0.2.126",
    base: "0.2.125",
    exitCode: 1,
    calls: ["metadata:head", "metadata:base"],
  },
])("$name", ({ head, base, exitCode, calls }) => {
  const directory = mkdtempSync(join(tmpdir(), "folio-footprint-workflow-"));
  directories.push(directory);
  const bin = join(directory, "bin");
  mkdirSync(bin);
  writeFileSync(join(directory, "Cargo.toml"), "head manifest");
  writeFileSync(join(directory, "cli-version"), head);
  writeFileSync(join(directory, "calls"), "");
  for (const name of ["cargo", "git", "wasm-bindgen", "bun", "jq"])
    writeFileSync(join(bin, name), `#!${process.execPath}\n${mockTool}`, { mode: 0o755 });
  const result = Bun.spawnSync(
    ["bash", "-euo", "pipefail", "-c", workflowStep("Compare named base and head projections")],
    {
      cwd: directory,
      env: {
        ...process.env,
        PATH: `${bin}:${process.env["PATH"] ?? ""}`,
        FOOTPRINT_FIXTURE: directory,
        RUNNER_TEMP: directory,
        BASE_SHA: "a".repeat(40),
        HEAD_SHA: "b".repeat(40),
        HEAD_VERSION: head,
        BASE_VERSION: base,
      },
      stdout: "pipe",
      stderr: "pipe",
    },
  );
  expect(result.exitCode, result.stderr.toString()).toBe(exitCode);
  expect(readFileSync(join(directory, "calls"), "utf8").trim().split("\n")).toEqual(calls);
  const report = join(directory, "projection-footprint.json");
  if (exitCode === 0) {
    expect(readFileSync(report, "utf8")).toBe("{}");
    return;
  }
  expect(result.stderr.toString()).toContain(
    `footprint attribution requires matching wasm-bindgen versions (head: ${head}, base: ${base})`,
  );
  expect(existsSync(report)).toBe(false);
});

test.each([
  {
    name: "raw Wasm budget failure",
    message: "Panic: DOCX kernel WebAssembly is 330835 bytes; budget is 318976",
    exitCode: 1,
    output: "budget_exceeded=true\n",
  },
  {
    name: "Brotli budget failure",
    message: "Panic: DOCX kernel Brotli size is 150000 bytes; budget is 149504",
    exitCode: 1,
    output: "budget_exceeded=true\n",
  },
  {
    name: "unrelated compilation failure",
    message: "error[E0308]: mismatched types",
    exitCode: 1,
    output: "",
  },
  {
    name: "non-default generator failure",
    message: "cargo invocation failed before artifact generation",
    exitCode: 17,
    output: "",
  },
  {
    name: "successful generation",
    message: "Generated canonical projection artifact",
    exitCode: 0,
    output: "",
  },
])(
  "generation preserves $name status and signals only budget failures",
  ({ message, exitCode, output }) => {
    const directory = mkdtempSync(join(tmpdir(), "folio-footprint-budget-"));
    directories.push(directory);
    const bin = join(directory, "bin");
    const githubOutput = join(directory, "github-output");
    mkdirSync(bin);
    writeFileSync(githubOutput, "");
    writeFileSync(
      join(bin, "bun"),
      `#!${process.execPath}\nconsole.error(process.env.BUILD_MESSAGE);\nprocess.exit(Number(process.env.BUILD_STATUS));\n`,
      { mode: 0o755 },
    );
    const result = Bun.spawnSync(
      ["bash", "-euo", "pipefail", "-c", workflowStep("Generate canonical projection artifact")],
      {
        cwd: directory,
        env: {
          ...process.env,
          PATH: `${bin}:${process.env["PATH"] ?? ""}`,
          RUNNER_TEMP: directory,
          GITHUB_OUTPUT: githubOutput,
          BUILD_MESSAGE: message,
          BUILD_STATUS: String(exitCode),
        },
        stdout: "pipe",
        stderr: "pipe",
      },
    );
    expect(result.exitCode, result.stderr.toString()).toBe(exitCode);
    expect(readFileSync(githubOutput, "utf8")).toBe(output);
    expect(readFileSync(join(directory, "canonical-projection-build.log"), "utf8")).toContain(
      message,
    );
  },
);

test("budget failure automatically schedules attribution even when the canonical job failed", () => {
  const jobs = workflowJobs();
  const canonical = jobs["canonical-wasm"];
  const footprint = jobs["footprint"];
  if (!isRecord(canonical) || !isRecord(canonical["outputs"]) || !isRecord(footprint))
    throw new TypeError("Canonical artifact and footprint jobs must exist.");
  expect(canonical["outputs"]["budget_exceeded"]).toBe(
    "${{ steps.generate.outputs.budget_exceeded }}",
  );
  expect(footprint["needs"]).toBe("canonical-wasm");
  expect(footprint["if"]).toBe(
    "always() && (needs.canonical-wasm.outputs.budget_exceeded == 'true' || (github.event_name == 'workflow_dispatch' && inputs.footprint_base != ''))",
  );
});
