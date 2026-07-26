#!/usr/bin/env bun

import { panic } from "better-result";
import { createHash } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { build } from "tsdown";

import agentsConfigs from "../packages/agents/tsdown.config";
import coreConfigs from "../packages/core/tsdown.config";

const BUILD_COUNT = 5;
const ROOT = join(import.meta.dirname, "..");
const targets = [
  {
    name: "@stll/folio-core",
    directory: join(ROOT, "packages/core"),
    config: coreConfigs.at(0),
  },
  {
    name: "@stll/folio-agents",
    directory: join(ROOT, "packages/agents"),
    config: agentsConfigs.at(0),
  },
];

const hashDirectory = async (directory: string): Promise<string> => {
  const files = Array.from(
    new Bun.Glob("**/*").scanSync({ cwd: directory, onlyFiles: true }),
  ).sort();
  const hash = createHash("sha256");
  for (const file of files) {
    const content = new Uint8Array(await Bun.file(join(directory, file)).arrayBuffer());
    hash.update(`${file.length}:`);
    hash.update(file);
    hash.update(`${content.byteLength}:`);
    hash.update(content);
  }
  return hash.digest("hex");
};

const tempRoot = await mkdtemp(join(tmpdir(), "folio-reproducible-build-"));
const originalDirectory = process.cwd();

try {
  for (const { name, directory, config } of targets) {
    if (!config) {
      panic(`Missing JavaScript build config for ${name}`);
    }

    process.chdir(directory);
    const digests: string[] = [];
    for (let index = 0; index < BUILD_COUNT; index += 1) {
      const outDir = join(tempRoot, name.replaceAll("/", "-"), String(index));
      await build({ ...config, outDir, clean: true, logLevel: "silent" });
      digests.push(await hashDirectory(outDir));
    }

    if (new Set(digests).size !== 1) {
      panic(`${name} emitted non-reproducible JavaScript: ${digests.join(", ")}`);
    }
    console.log(`${name}: ${BUILD_COUNT} byte-identical JavaScript builds`);
  }
} finally {
  process.chdir(originalDirectory);
  await rm(tempRoot, { recursive: true, force: true });
}
