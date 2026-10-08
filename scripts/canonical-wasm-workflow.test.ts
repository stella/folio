import { panic } from "better-result";
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const repoRoot = path.resolve(import.meta.dir, "..");

test("canonical artifact triggers cover their package command owners", () => {
  const workflow: unknown = Bun.YAML.parse(
    readFileSync(path.join(repoRoot, ".github/workflows/canonical-wasm.yml"), "utf8"),
  );
  if (!isRecord(workflow) || !isRecord(workflow["on"]) || !isRecord(workflow["jobs"])) {
    return panic("Invalid canonical artifact workflow");
  }
  const trigger = workflow["on"]["pull_request"];
  if (!isRecord(trigger) || !Array.isArray(trigger["paths"])) {
    return panic("Canonical artifact workflow needs PR path filters");
  }
  const paths = trigger["paths"];
  if (!paths.every((pattern: unknown) => typeof pattern === "string")) {
    return panic("Canonical artifact path filters must be strings");
  }
  const selected = new Set<string>();
  for (const job of Object.values(workflow["jobs"])) {
    if (!isRecord(job) || !Array.isArray(job["steps"])) continue;
    for (const step of job["steps"]) {
      if (!isRecord(step) || typeof step["run"] !== "string") continue;
      const name = step["run"].match(/bun --filter '([^']+)' wasm:generate/u)?.at(1);
      if (name !== undefined) selected.add(name);
    }
  }
  expect(selected.size).toBeGreaterThan(0);
  const exercised = new Set<string>();
  for (const manifestPath of new Bun.Glob("packages/*/package.json").scanSync({ cwd: repoRoot })) {
    const manifest: unknown = JSON.parse(readFileSync(path.join(repoRoot, manifestPath), "utf8"));
    if (!isRecord(manifest) || typeof manifest["name"] !== "string") {
      return panic("Invalid workspace manifest");
    }
    if (!selected.has(manifest["name"])) continue;
    if (
      !isRecord(manifest["scripts"]) ||
      typeof manifest["scripts"]["wasm:generate"] !== "string"
    ) {
      return panic("Selected package needs a WASM generation command");
    }
    const script = manifest["scripts"]["wasm:generate"].match(/bun ([^\s]+\.ts)(?:\s|$)/u)?.at(1);
    if (script === undefined) return panic("Selected WASM command needs a generator script");
    const generatorPath = path.normalize(path.join(path.dirname(manifestPath), script));
    for (const input of [manifestPath, generatorPath]) {
      expect(
        paths.some((pattern) => new Bun.Glob(pattern).match(input)),
        input,
      ).toBe(true);
    }
    exercised.add(manifest["name"]);
  }
  expect(exercised).toEqual(selected);
});
