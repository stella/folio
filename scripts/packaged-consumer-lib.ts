// Shared building blocks for the packaged-consumer gates.
//
// Both the build-only gate (`packaged-consumer-build.ts`) and the runtime smoke
// (`packaged-consumer-smoke.ts`) start from the same clean-room recipe: build
// both @stll/folio-* packages to their published "dist" shape, pack a tarball
// each (`bun pm pack`, which pins workspace/catalog protocols), and stage a
// consumer OUTSIDE the monorepo so no workspace linkage leaks in. Only the last
// mile differs — the build-only gate runs a production `vite build` and inspects
// dist; the smoke additionally serves that build and drives it in a browser.
//
// Keeping the pack + consumer-manifest logic here means the subtle bits (the
// reversible package.json transform, the `overrides` pin that resolves
// folio-react's transitive folio-core to the packed tarball) live in one place.

import { panic } from "better-result";
import { $ } from "bun";
import { readFile, readdir, writeFile } from "node:fs/promises";
import path from "node:path";

export const repoRoot = path.resolve(import.meta.dir, "..");
export const prepareScript = path.join(repoRoot, "scripts", "prepare-publish.ts");
export const coreDir = path.join(repoRoot, "packages", "core");
export const reactDir = path.join(repoRoot, "packages", "react");
export const consumerSrc = path.join(repoRoot, "test", "packaged-consumer");

// Build a package, transform its package.json to the published dist shape
// (reversibly), pack a tarball into destDir, and return the tarball path.
export const buildAndPack = async (pkgDir: string, destDir: string): Promise<string> => {
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

// Write the clean-room consumer's package.json. The `overrides` entry pins
// folio-react's transitive @stll/folio-core to the packed tarball (not a
// registry lookup): the two must resolve to each other's packed dist.
export const writeConsumerPackageJson = async (
  consumerDir: string,
  coreTarball: string,
): Promise<void> => {
  const consumerPkg = {
    name: "folio-packaged-consumer",
    version: "0.0.0",
    private: true,
    type: "module",
    overrides: { "@stll/folio-core": coreTarball },
  };
  await writeFile(
    path.join(consumerDir, "package.json"),
    `${JSON.stringify(consumerPkg, null, 2)}\n`,
  );
};
