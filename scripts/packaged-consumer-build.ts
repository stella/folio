#!/usr/bin/env bun
// Clean-room packaging gate: does a production Vite build of a real downstream
// consumer succeed against the PACKED @stll/folio-* tarballs?
//
// The in-repo playground (tests/visual, `bun run test:interactions`) consumes
// folio from SOURCE, so its bundler resolves `./src/*.ts` directly and never
// sees the published dist. That is exactly why it cannot catch a packaging
// regression such as a worker `new URL("...worker.ts", import.meta.url)` target
// that the dist build never emits: source has the `.ts`, dist does not. This
// gate closes that gap.
//
// Steps (mirrors scripts/validate-dist.ts's clean room):
//   1. Build + transform both packages to their published "dist" shape and pack
//      a tarball each (`bun pm pack`, which pins workspace/catalog protocols).
//   2. Copy test/packaged-consumer OUT of the monorepo, install the tarballs
//      into it npm-style (an `overrides` pin resolves folio-react's transitive
//      folio-core to the packed tarball, since neither is on a registry).
//   3. Run `vite build`. The `vite:worker-import-meta-url` plugin turns every
//      `new URL(..., import.meta.url)` worker into its own build entry; a target
//      missing from the published dist aborts with UNRESOLVED_ENTRY.
//   4. Assert the build emitted the worker chunk, proving the worker URL
//      resolved end to end.
//
// Exits non-zero on any failure. Run via `bun run test:packaged-consumer`.

import { panic } from "better-result";
import { $ } from "bun";
import { cp, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dir, "..");
const prepareScript = path.join(repoRoot, "scripts", "prepare-publish.ts");
const coreDir = path.join(repoRoot, "packages", "core");
const reactDir = path.join(repoRoot, "packages", "react");
const consumerSrc = path.join(repoRoot, "test", "packaged-consumer");

// Build a package, transform its package.json to the published dist shape
// (reversibly), pack a tarball into destDir, and return the tarball path.
const buildAndPack = async (pkgDir: string, destDir: string): Promise<string> => {
  const name = path.basename(pkgDir);
  console.log(`→ building @stll/folio-${name}`);
  await $`bun run build`.cwd(pkgDir).quiet();

  const pkgJsonPath = path.join(pkgDir, "package.json");
  const original = await readFile(pkgJsonPath, "utf-8");
  try {
    console.log(`→ transforming @stll/folio-${name} package.json to dist shape`);
    await $`bun ${prepareScript} ${pkgDir}`.quiet();
    console.log(`→ packing @stll/folio-${name} tarball`);
    await $`bun pm pack --destination ${destDir}`.cwd(pkgDir).quiet();
  } finally {
    // Always restore the in-repo source-shape package.json.
    await writeFile(pkgJsonPath, original);
  }
  const tgz = (await readdir(destDir)).find((f) => f.endsWith(".tgz"));
  return tgz
    ? path.join(destDir, tgz)
    : panic(`packaged-consumer: bun pm pack produced no tarball for ${name}`);
};

const corePackDir = await mkdtemp(path.join(tmpdir(), "folio-core-pack-"));
const reactPackDir = await mkdtemp(path.join(tmpdir(), "folio-react-pack-"));
const consumerDir = await mkdtemp(path.join(tmpdir(), "folio-packaged-consumer-"));

const coreTarball = await buildAndPack(coreDir, corePackDir);
const reactTarball = await buildAndPack(reactDir, reactPackDir);

// Copy the checked-in consumer app out of the monorepo so no workspace linkage
// leaks in; install the packed artifacts npm-style.
console.log(`→ staging consumer in ${consumerDir}`);
await cp(path.join(consumerSrc, "src"), path.join(consumerDir, "src"), { recursive: true });
await cp(path.join(consumerSrc, "vite.config.ts"), path.join(consumerDir, "vite.config.ts"));

const consumerPkg = {
  name: "folio-packaged-consumer",
  version: "0.0.0",
  private: true,
  type: "module",
  // Pin folio-react's transitive folio-core to the packed tarball (not a
  // registry lookup): the two must resolve to each other's packed dist.
  overrides: { "@stll/folio-core": coreTarball },
};
await writeFile(
  path.join(consumerDir, "package.json"),
  `${JSON.stringify(consumerPkg, null, 2)}\n`,
);

console.log("→ installing tarballs + peers");
await $`bun add ${reactTarball} ${coreTarball} react@^19 react-dom@^19 use-intl@^4 vite@^8 @types/react@^19 @types/react-dom@^19`
  .cwd(consumerDir)
  .quiet();

console.log("→ production vite build over the packed tarballs");
const build = await $`bun x vite build`.cwd(consumerDir).nothrow();
if (build.exitCode !== 0) {
  console.error(build.stderr.toString() || build.stdout.toString());
  console.error("\n✗ packaged-consumer: production vite build FAILED against the packed tarballs.");
  process.exit(1);
}

// The worker must have resolved and emitted its own chunk; otherwise the URL
// target silently vanished (e.g. externalized away) and the gate would be moot.
const distFiles = await readdir(path.join(consumerDir, "dist"), { recursive: true });
const workerChunk = distFiles.find((f) => f.includes("font-metrics.worker") && f.endsWith(".js"));
if (!workerChunk) {
  console.error(
    `✗ packaged-consumer: build produced no worker chunk. dist: ${distFiles.join(", ")}`,
  );
  process.exit(1);
}

await rm(corePackDir, { recursive: true, force: true });
await rm(reactPackDir, { recursive: true, force: true });
await rm(consumerDir, { recursive: true, force: true });

console.log(
  `\n✓ packaged-consumer: production vite build succeeded; worker chunk emitted (${workerChunk}).`,
);
