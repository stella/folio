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

const comparisonStep = () => {
  const workflow: unknown = Bun.YAML.parse(
    readFileSync(join(import.meta.dir, "../.github/workflows/canonical-wasm.yml"), "utf8"),
  );
  if (!isRecord(workflow) || !isRecord(workflow["jobs"]))
    throw new TypeError("Canonical workflow has no jobs.");
  for (const job of Object.values(workflow["jobs"])) {
    if (!isRecord(job) || !Array.isArray(job["steps"])) continue;
    for (const step of job["steps"]) {
      if (!isRecord(step) || step["name"] !== "Compare named base and head projections") continue;
      if (typeof step["run"] !== "string")
        throw new TypeError("Projection comparison step has no executable script.");
      return step["run"];
    }
  }
  throw new TypeError("Canonical workflow has no projection comparison step.");
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
  const result = Bun.spawnSync(["bash", "-euo", "pipefail", "-c", comparisonStep()], {
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
  });
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
