// Shared building blocks for the packaged-consumer gates.
//
// Both the build-only gate (`packaged-consumer-build.ts`) and the runtime smoke
// (`packaged-consumer-smoke.ts`) start from the same clean-room recipe: build
// the workspace packages to their published "dist" shape, pack a tarball each
// (`bun pm pack`, which pins workspace/catalog protocols), and stage a
// consumer OUTSIDE the monorepo so no workspace linkage leaks in. Only the last
// mile differs — the build-only gate runs a production `vite build` and inspects
// dist; the smoke additionally serves that build and drives it in a browser.
//
// Keeping the pack + consumer-manifest logic here means the subtle bits (the
// reversible package.json transform, the `overrides` pin that resolves
// folio-react's transitive folio-core to the packed tarball) live in one place.

import { panic } from "better-result";
import reactManifest from "../packages/react/package.json";
import coreManifest from "../packages/core/package.json";
import { $ } from "bun";
import { readFile, readdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { assertHostPeersExternal } from "./host-peer-policy";

export const repoRoot = path.resolve(import.meta.dir, "..");
export const prepareScript = path.join(repoRoot, "scripts", "prepare-publish.ts");
export const docxCoreDir = path.join(repoRoot, "packages", "docx-core");
export const utf8BoundariesDir = path.join(repoRoot, "packages", "utf8-boundaries");
export const coreDir = path.join(repoRoot, "packages", "core");
export const reactDir = path.join(repoRoot, "packages", "react");
export const consumerSrc = path.join(repoRoot, "test", "packaged-consumer");

type ReactPeers = { react: string; "react-dom": string };

export const reactPeerFloor = (peers: ReactPeers) => {
  if (peers.react !== peers["react-dom"]) {
    panic("React and React DOM peer ranges must match");
  }
  const match = /^\^(\d+)\.\d+\.\d+$/u.exec(peers.react);
  if (!match) panic("Consumer checks require one caret React peer range");
  return {
    version: peers.react.slice(1),
    major: match.at(1) ?? panic("React peer range must include a major version"),
  };
};

export const REACT_PEER_FLOOR = reactPeerFloor(reactManifest.peerDependencies);

const DOCX_KERNEL_WASM_ASSET = /(?:^|\/)docx_kernel_bg(?:-[^/]+)?\.wasm$/u;

export const isDocxKernelWasmAsset = (file: string): boolean => DOCX_KERNEL_WASM_ASSET.test(file);

export const reactPeerInstallArgs = () => [
  `react@${REACT_PEER_FLOOR.version}`,
  `react-dom@${REACT_PEER_FLOOR.version}`,
  `@types/react@^${REACT_PEER_FLOOR.major}`,
  `@types/react-dom@^${REACT_PEER_FLOOR.major}`,
];

// The host installs the runtime peers itself; no singleton override or alias.
export const hostPeerInstallArgs = () =>
  Object.entries(coreManifest.peerDependencies).map(([name, range]) => `${name}@${range}`);

// Build a package, transform its package.json to the published dist shape
// (reversibly), pack a tarball into destDir, and return the tarball path.
export const buildAndPack = async (pkgDir: string, destDir: string): Promise<string> => {
  const name = path.basename(pkgDir);
  const packageJson = (await Bun.file(path.join(pkgDir, "package.json")).json()) as {
    name: string;
  };
  const sharedTarballs = process.env["FOLIO_PACKED_TARBALLS"];
  if (sharedTarballs) {
    const tarballs = await Bun.file(path.join(sharedTarballs, "tarballs.json")).json();
    const tarball = tarballs[packageJson.name];
    if (typeof tarball !== "string" || !(await Bun.file(tarball).exists()))
      panic(`packaged-consumer: no shared tarball for ${packageJson.name}`);
    return tarball;
  }
  console.log(`→ building ${packageJson.name}`);
  await $`bun run build`.cwd(pkgDir).quiet();
  assertHostPeersExternal(pkgDir);

  const pkgJsonPath = path.join(pkgDir, "package.json");
  const original = await readFile(pkgJsonPath, "utf-8");
  try {
    const manifest = JSON.parse(original);
    if (
      Object.values(manifest.exports).some(
        (target) => typeof target === "string" && target.startsWith("./src/"),
      )
    ) {
      console.log(`→ transforming ${packageJson.name} package.json to dist shape`);
      await $`bun ${prepareScript} ${pkgDir}`.quiet();
    }
    console.log(`→ packing ${packageJson.name} tarball`);
    await $`bun pm pack --ignore-scripts --destination ${destDir}`.cwd(pkgDir).quiet();
  } finally {
    // Always restore the in-repo source-shape package.json.
    await writeFile(pkgJsonPath, original);
  }
  const tgz = (await readdir(destDir)).find((f) => f.endsWith(".tgz"));
  return tgz
    ? path.join(destDir, tgz)
    : panic(`packaged-consumer: bun pm pack produced no tarball for ${name}`);
};

type WriteConsumerPackageJsonOptions = {
  consumerDir: string;
  coreTarball: string;
  docxCoreTarball: string;
  utf8BoundariesTarball: string;
};

// Write the clean-room consumer's package.json. The `overrides` entries pin
// transitive workspace dependencies to the coordinated packed tarballs.
export const writeConsumerPackageJson = async ({
  consumerDir,
  coreTarball,
  docxCoreTarball,
  utf8BoundariesTarball,
}: WriteConsumerPackageJsonOptions): Promise<void> => {
  const consumerPkg = {
    name: "folio-packaged-consumer",
    version: "0.0.0",
    private: true,
    type: "module",
    overrides: {
      "@stll/docx-core": docxCoreTarball,
      "@stll/folio-core": coreTarball,
      "@stll/utf8-boundaries": utf8BoundariesTarball,
    },
  };
  await writeFile(
    path.join(consumerDir, "package.json"),
    `${JSON.stringify(consumerPkg, null, 2)}\n`,
  );
};
