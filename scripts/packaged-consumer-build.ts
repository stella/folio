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
//   1. Build + transform the packages to their published "dist" shape and pack
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

import { $ } from "bun";
import { cp, mkdtemp, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  REACT_PEER_MAJORS,
  buildAndPack,
  consumerSrc,
  coreDir,
  docxCoreDir,
  reactPeerInstallArgs,
  reactDir,
  type ReactPeerMajor,
  writeConsumerPackageJson,
} from "./packaged-consumer-lib";

// Track the failure instead of calling `process.exit` inside the `try`:
// `process.exit` terminates without unwinding, so the `finally` cleanup would
// be skipped — the exact leak this structure exists to prevent. Any thrown
// error still propagates (Bun exits non-zero on it) after cleanup runs.
let failure: string | null = null;
const workerChunks: string[] = [];

let docxCorePackDir = "";
let corePackDir = "";
let reactPackDir = "";
const consumerDirs: string[] = [];
try {
  docxCorePackDir = await mkdtemp(path.join(tmpdir(), "folio-docx-core-pack-"));
  corePackDir = await mkdtemp(path.join(tmpdir(), "folio-core-pack-"));
  reactPackDir = await mkdtemp(path.join(tmpdir(), "folio-react-pack-"));

  const docxCoreTarball = await buildAndPack(docxCoreDir, docxCorePackDir);
  const coreTarball = await buildAndPack(coreDir, corePackDir);
  const reactTarball = await buildAndPack(reactDir, reactPackDir);

  const runBuild = async (reactMajor: ReactPeerMajor): Promise<void> => {
    const consumerDir = await mkdtemp(
      path.join(tmpdir(), `folio-packaged-consumer-r${reactMajor}-`),
    );
    consumerDirs.push(consumerDir);

    // Copy the checked-in consumer app out of the monorepo so no workspace
    // linkage leaks in; install the packed artifacts npm-style.
    console.log(`→ staging React ${reactMajor} consumer in ${consumerDir}`);
    await cp(path.join(consumerSrc, "src"), path.join(consumerDir, "src"), { recursive: true });
    await cp(path.join(consumerSrc, "vite.config.ts"), path.join(consumerDir, "vite.config.ts"));

    await writeConsumerPackageJson({ consumerDir, coreTarball, docxCoreTarball });

    console.log(`→ installing tarballs + React ${reactMajor} peers`);
    await $`bun add ${reactTarball} ${coreTarball} ${docxCoreTarball} ${reactPeerInstallArgs(reactMajor)} use-intl@^4 vite@^8`
      .cwd(consumerDir)
      .quiet();

    console.log(`→ production vite build over the packed tarballs (React ${reactMajor})`);
    const build = await $`bun x vite build`.cwd(consumerDir).nothrow();
    if (build.exitCode !== 0) {
      console.error(build.stderr.toString() || build.stdout.toString());
      failure = `\n✗ packaged-consumer: React ${reactMajor} production vite build FAILED against the packed tarballs.`;
      return;
    }

    // The worker must have resolved and emitted its own chunk; otherwise the
    // URL target silently vanished (e.g. externalized away) and the gate would
    // be moot.
    const distFiles = await readdir(path.join(consumerDir, "dist"), { recursive: true });
    const workerChunk = distFiles.find(
      (f) => f.includes("font-metrics.worker") && f.endsWith(".js"),
    );
    if (!workerChunk) {
      failure = `✗ packaged-consumer: React ${reactMajor} build produced no worker chunk. dist: ${distFiles.join(", ")}`;
      return;
    }
    workerChunks.push(`React ${reactMajor}: ${workerChunk}`);
  };

  for (const reactMajor of REACT_PEER_MAJORS) {
    await runBuild(reactMajor);
    if (failure !== null) {
      break;
    }
  }
} finally {
  // `rm` with `force` tolerates a dir that was never created (empty string is
  // guarded); cleanup runs on success, on failure, and on a thrown step.
  for (const dir of [docxCorePackDir, corePackDir, reactPackDir, ...consumerDirs]) {
    if (dir !== "") {
      await rm(dir, { recursive: true, force: true });
    }
  }
}

if (failure !== null) {
  console.error(failure);
  process.exit(1);
}

console.log(
  `\n✓ packaged-consumer: production vite build succeeded for React 18 and 19; worker chunks emitted (${workerChunks.join(", ")}).`,
);
