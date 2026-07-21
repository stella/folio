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
// The only fix once it drifts is `rm bun.lock && bun install` (a full
// regenerate). This script cross-checks each packages/*/package.json
// `version` against the version bun.lock has cached for that workspace, so
// the drift itself gets caught in CI instead of silently persisting.

import { readdir } from "node:fs/promises";
import { join } from "node:path";

const ROOT = new URL("..", import.meta.url).pathname;

const readJson = async (path: string): Promise<Record<string, unknown>> =>
  JSON.parse(await Bun.file(path).text());

const packagesDir = join(ROOT, "packages");
const entries = await readdir(packagesDir, { withFileTypes: true });
const workspaceDirs = entries
  .filter((entry) => entry.isDirectory())
  .map((entry) => `packages/${entry.name}`)
  .sort();

const lockText = await Bun.file(join(ROOT, "bun.lock")).text();

// bun.lock is JSON-with-trailing-commas ("JSONC"-flavored), not strict JSON,
// so a plain JSON.parse fails on it. Rather than hand-roll a tolerant parser
// for the whole file (risking mis-parsing the many base64 `sha512-...`
// strings that legitimately contain `//`), extract just the one shape we
// need directly: the `"<workspace path>": { "name": "...", "version": "..." }`
// block bun emits per workspace entry.
const versionForWorkspace = (workspacePath: string): string | null => {
  const escaped = workspacePath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const blockRe = new RegExp(`"${escaped}":\\s*\\{([^}]*)\\}`);
  const block = blockRe.exec(lockText)?.[1];
  if (!block) return null;
  return /"version":\s*"([^"]+)"/.exec(block)?.[1] ?? null;
};

const mismatches: string[] = [];

for (const workspaceDir of workspaceDirs) {
  const pkg = await readJson(join(ROOT, workspaceDir, "package.json"));
  const name = pkg.name;
  const version = pkg.version;
  if (typeof name !== "string" || typeof version !== "string") continue;

  const lockedVersion = versionForWorkspace(workspaceDir);
  if (lockedVersion === null) {
    mismatches.push(`${name} (${workspaceDir}): no bun.lock entry found`);
    continue;
  }
  if (lockedVersion !== version) {
    mismatches.push(
      `${name} (${workspaceDir}): package.json is ${version}, bun.lock has ${lockedVersion}`,
    );
  }
}

if (mismatches.length > 0) {
  console.error(
    [
      "bun.lock workspace-version drift detected:",
      "",
      ...mismatches.map((line) => `  - ${line}`),
      "",
      "A plain `bun install` will not fix this (it doesn't rewrite cached",
      "workspace versions for entries that already exist). Regenerate the",
      "lockfile instead:",
      "",
      "    rm bun.lock && bun install",
      "",
      "Then commit the refreshed bun.lock.",
    ].join("\n"),
  );
  process.exit(1);
}

console.log("bun.lock workspace-version check: all workspace versions match. OK.");
