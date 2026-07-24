#!/usr/bin/env bun
// CI gate: catches stale workspace `"version"` fields cached in bun.lock.
//
// Why this exists: `bun install` (even non-frozen) does NOT rewrite the
// `"version"` field bun.lock records for an already-present workspace entry
// when only that package's own package.json version changed — it only
// re-resolves dependency ranges. `bun install --frozen-lockfile` (what CI
// runs everywhere) validates that the dependency graph still satisfies the
// lockfile; it does not compare workspace self-versions either. So neither
// the normal install path nor the frozen-lockfile CI gate ever notices a
// workspace's recorded version drifting behind its package.json — and
// `bun pm pack` reads the *lockfile's* cached version when resolving
// workspace:* ranges for publish, so a stale entry silently ships a wrong
// dependency range (see the @stll/docx-core ^0.4.0-vs-0.5.0 incident this
// script was added to prevent).
//
// This script is the single owner of workspace self-version synchronization:
// check-only by default for CI, or byte-preserving repair with `--write`.

import { panic } from "better-result";
import { readdir } from "node:fs/promises";
import { join } from "node:path";
import { syncWorkspaceVersions } from "./lib/bun-lock-workspace-versions";

const ROOT = join(import.meta.dirname, "..");

const readJson = async (path: string): Promise<Record<string, unknown>> =>
  JSON.parse(await Bun.file(path).text());

const packagesDir = join(ROOT, "packages");
const entries = await readdir(packagesDir, { withFileTypes: true });
const workspaceDirs = entries
  .filter((entry) => entry.isDirectory())
  .map((entry) => `packages/${entry.name}`)
  .sort();

const lockText = await Bun.file(join(ROOT, "bun.lock")).text();

const args = process.argv.slice(2);
if (args.some((arg) => arg !== "--write") || args.filter((arg) => arg === "--write").length > 1) {
  panic("Usage: bun scripts/check-lockfile-workspace-versions.ts [--write]");
}
const write = args[0] === "--write";
const expectedVersions = new Map<string, string>();
const packageNames = new Map<string, string>();

for (const workspaceDir of workspaceDirs) {
  // A directory under packages/ is not necessarily a real workspace: skip
  // it (rather than crash the guard) if its package.json is missing or
  // fails to parse.
  const pkg = await readJson(join(ROOT, workspaceDir, "package.json")).catch(() => null);
  if (pkg === null) continue;
  const name = pkg.name;
  const version = pkg.version;
  if (typeof name !== "string" || typeof version !== "string") continue;

  expectedVersions.set(workspaceDir, version);
  packageNames.set(workspaceDir, name);
}

const result = syncWorkspaceVersions(lockText, expectedVersions);
const unrepairable = result.mismatches.filter(({ actual }) => actual === null);

if (write && unrepairable.length === 0 && result.text !== lockText) {
  await Bun.write(join(ROOT, "bun.lock"), result.text);
  console.log(`bun.lock workspace-version sync: updated ${result.mismatches.length} workspace(s).`);
  process.exit(0);
}

const mismatches = result.mismatches.map(({ workspace, expected, actual }) =>
  actual === null
    ? `${packageNames.get(workspace)} (${workspace}): no writable bun.lock version entry found`
    : `${packageNames.get(workspace)} (${workspace}): package.json is ${expected}, bun.lock has ${actual}`,
);

if (mismatches.length > 0) {
  console.error(
    [
      "bun.lock workspace-version drift detected:",
      "",
      ...mismatches.map((line) => `  - ${line}`),
      "",
      write
        ? "The lockfile shape is incomplete; workspace entries must exist before they can be synchronized."
        : "A plain `bun install` will not fix cached workspace self-versions. Synchronize them with:",
      "",
      "    bun scripts/check-lockfile-workspace-versions.ts --write",
      "    bun install --frozen-lockfile",
      "",
      "Then commit the refreshed bun.lock.",
    ].join("\n"),
  );
  process.exit(1);
}

console.log(
  write
    ? "bun.lock workspace-version sync: already current. OK."
    : "bun.lock workspace-version check: all workspace versions match. OK.",
);
