#!/usr/bin/env bun

import { panic } from "better-result";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const write = process.argv[2] === "--write";
if (process.argv.length > 3 || (process.argv[2] !== undefined && !write)) {
  panic("usage: bun scripts/sync-docx-kernel-version.ts [--write]");
}

const repoRoot = path.resolve(import.meta.dir, "..");
const packageJsonPath = path.join(repoRoot, "packages", "docx-core", "package.json");
const cargoManifestPath = path.join(repoRoot, "Cargo.toml");
const packageJson = (await Bun.file(packageJsonPath).json()) as { version?: unknown };
if (typeof packageJson.version !== "string") {
  panic("@stll/docx-core package.json has no string version");
}

const manifest = await readFile(cargoManifestPath, "utf8");
const manifestVersionPattern = /(\[workspace\.package\][\s\S]*?\nversion = ")([^"]+)(")/u;
const manifestMatch = manifest.match(manifestVersionPattern);
const manifestVersion = manifestMatch?.[2] ?? panic("Cargo workspace package version is missing");

if (write) {
  await writeFile(
    cargoManifestPath,
    manifest.replace(
      manifestVersionPattern,
      (_match, prefix: string, _version: string, suffix: string) =>
        `${prefix}${packageJson.version}${suffix}`,
    ),
  );
  // Cargo owns the lock graph, including every crate inheriting the workspace version.
  const update = Bun.spawnSync(["cargo", "update", "--workspace"], {
    cwd: repoRoot,
    stdout: "inherit",
    stderr: "inherit",
  });
  if (update.exitCode !== 0) panic("Failed to synchronize Cargo workspace lock versions");
  console.log(`Synchronized DOCX kernel to ${packageJson.version}`);
} else {
  if (manifestVersion !== packageJson.version) {
    panic(`DOCX kernel version drift: npm=${packageJson.version}, Cargo.toml=${manifestVersion}`);
  }
  const metadata = Bun.spawnSync(["cargo", "metadata", "--locked", "--format-version", "1"], {
    cwd: repoRoot,
    stdout: "ignore",
    stderr: "inherit",
  });
  if (metadata.exitCode !== 0) panic("Cargo.lock is not synchronized with the workspace");
  console.log(`DOCX kernel version ${packageJson.version} is synchronized`);
}
