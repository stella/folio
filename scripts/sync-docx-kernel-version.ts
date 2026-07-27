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
const cargoLockPath = path.join(repoRoot, "Cargo.lock");
const packageJson = (await Bun.file(packageJsonPath).json()) as { version?: unknown };
if (typeof packageJson.version !== "string") {
  panic("@stll/docx-core package.json has no string version");
}

const manifest = await readFile(cargoManifestPath, "utf8");
const manifestVersionPattern = /(\[workspace\.package\][\s\S]*?\nversion = ")([^"]+)(")/u;
const manifestMatch = manifest.match(manifestVersionPattern);
const manifestVersion = manifestMatch?.[2] ?? panic("Cargo workspace package version is missing");

const lock = await readFile(cargoLockPath, "utf8");
const lockVersionPattern = /(\[\[package\]\]\nname = "stella-docx-kernel"\nversion = ")([^"]+)(")/u;
const lockMatch = lock.match(lockVersionPattern);
const lockVersion = lockMatch?.[2] ?? panic("stella-docx-kernel Cargo.lock entry is missing");

if (write) {
  await writeFile(
    cargoManifestPath,
    manifest.replace(
      manifestVersionPattern,
      (_match, prefix: string, _version: string, suffix: string) =>
        `${prefix}${packageJson.version}${suffix}`,
    ),
  );
  await writeFile(
    cargoLockPath,
    lock.replace(
      lockVersionPattern,
      (_match, prefix: string, _version: string, suffix: string) =>
        `${prefix}${packageJson.version}${suffix}`,
    ),
  );
  console.log(`Synchronized DOCX kernel to ${packageJson.version}`);
} else {
  if (manifestVersion !== packageJson.version || lockVersion !== packageJson.version) {
    panic(
      `DOCX kernel version drift: npm=${packageJson.version}, Cargo.toml=${manifestVersion}, Cargo.lock=${lockVersion}`,
    );
  }
  console.log(`DOCX kernel version ${packageJson.version} is synchronized`);
}
