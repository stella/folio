#!/usr/bin/env bun
// All published packages use the existing clean-room pack recipe, once per run.
import { $ } from "bun";
import { mkdir, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { buildAndPack, repoRoot } from "./packaged-consumer-lib";
import { installCleanRoomDependencies } from "./clean-room-install";
import { PUBLISHED_PACKAGES } from "./lib/published-packages";
import { publicEntries } from "./lib/packed-exports";

const packIndex = process.argv.indexOf("--pack-dir");
const packDir =
  packIndex < 0
    ? await mkdtemp(path.join(tmpdir(), "folio-dependency-packs-"))
    : path.resolve(process.argv.at(packIndex + 1) ?? "test-results/package-tarballs");
await mkdir(packDir, { recursive: true });
const tarballs: Record<string, string> = {};
try {
  for (const pkg of PUBLISHED_PACKAGES) {
    const destination = path.join(packDir, pkg.slug);
    await mkdir(destination, { recursive: true });
    tarballs[pkg.name] = await buildAndPack(pkg.root, destination);
  }
  await writeFile(path.join(packDir, "tarballs.json"), JSON.stringify(tarballs));
  // The same rules check emitted JS and declarations after the shared build.
  await $`bun scripts/check-package-dependencies.ts --dist`.cwd(repoRoot);

  for (const pkg of PUBLISHED_PACKAGES) {
    const consumer = await mkdtemp(path.join(tmpdir(), "folio-dependency-consumer-"));
    try {
      await writeFile(
        path.join(consumer, "package.json"),
        JSON.stringify({
          name: "folio-dependency-consumer",
          private: true,
          type: "module",
          overrides: tarballs,
        }),
      );
      const archive = tarballs[pkg.name];
      if (!archive) throw new TypeError(`No tarball for ${pkg.name}`);
      // Read the actual packed manifest, including rewritten workspace versions.
      const unpacked = path.join(consumer, "packed");
      await mkdir(unpacked);
      await $`tar -xzf ${archive} -C ${unpacked}`.quiet();
      const packedRoot = path.join(unpacked, "package");
      const manifest = await Bun.file(path.join(packedRoot, "package.json")).json();
      const peers = Object.entries(manifest.peerDependencies ?? {}).map(([name, range]) => {
        if (typeof range !== "string") throw new TypeError(`Invalid peer range for ${name}`);
        return tarballs[name] ?? `${name}@${range}`;
      });
      await installCleanRoomDependencies({
        consumerDir: consumer,
        dependencies: [archive, ...peers],
        linker: "isolated",
      });
      // Keep package resolution isolated even if Bun's global config chooses hoisting.
      // No repository tool/runtime packages are added to the consumer.
      const entries = publicEntries({
        name: pkg.name,
        exports: manifest.exports,
        files: await readdir(packedRoot, { recursive: true }),
      });
      if (entries.length === 0) throw new TypeError(`No public entries for ${pkg.name}`);
      const esm = entries.filter((entry) => entry.esm).map((entry) => entry.specifier);
      const cjs = entries.filter((entry) => entry.cjs).map((entry) => entry.specifier);
      const assets = entries.filter((entry) => entry.asset).map((entry) => entry.specifier);
      await writeFile(
        path.join(consumer, "runtime.mjs"),
        `
import { createRequire } from 'node:module';
import { access } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
for (const entry of ${JSON.stringify(esm)}) { await import(entry); }
for (const entry of ${JSON.stringify(cjs)}) { require(entry); }
for (const entry of ${JSON.stringify(assets)}) { await access(fileURLToPath(import.meta.resolve(entry))); }
`,
      );
      await $`node runtime.mjs`.cwd(consumer);
      const typed = entries.filter((entry) => !entry.asset);
      await writeFile(
        path.join(consumer, "consumer.ts"),
        typed
          .map(
            (entry, index) =>
              `import * as entry${index} from ${JSON.stringify(entry.specifier)};\nvoid entry${index};`,
          )
          .join("\n"),
      );
      await writeFile(
        path.join(consumer, "tsconfig.json"),
        JSON.stringify({
          compilerOptions: {
            target: "ES2022",
            module: "Preserve",
            moduleResolution: "Bundler",
            strict: true,
            noEmit: true,
            skipLibCheck: true,
            types: [],
            lib: ["ES2022", "DOM", "DOM.Iterable"],
          },
          files: ["consumer.ts"],
        }),
      );
      const compiler = path.join(repoRoot, "node_modules", "typescript", "bin", "tsc");
      await $`node ${compiler} -p tsconfig.json`.cwd(consumer);
      console.log(
        `PASS ${pkg.name}: ${esm.length} ESM, ${cjs.length} CJS, ${typed.length} type entries`,
      );
    } finally {
      await rm(consumer, { recursive: true, force: true });
    }
  }
} finally {
  if (packIndex < 0) await rm(packDir, { recursive: true, force: true });
}
