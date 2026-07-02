#!/usr/bin/env bun
// Runtime half of the packaged-consumer gate: build + pack both @stll/folio-*
// packages, stage the clean-room consumer against the packed tarballs, run a
// production SPA build, SERVE it, and drive it in a real browser (Playwright).
//
// The build-only gate (`packaged-consumer-build.ts`) proves the tarballs RESOLVE
// under a production bundler. This proves the built output RUNS: it catches the
// two failure classes that stay invisible until a real consumer executes the
// code — a font-metrics worker that resolves at build time but fails to
// spawn/execute, and UI strings missing from the shipped runtime catalog
// (`IntlError: MISSING_MESSAGE`). Assertions live in
// `test/packaged-consumer/smoke.spec.ts`.
//
// This script owns the app lifecycle (build + `vite preview`) and hands the
// preview URL to Playwright via `PLAYWRIGHT_BASE_URL`; the browser install/cache
// is the same one the interaction-e2e job already sets up. Exits non-zero on any
// failure. Run via `bun run test:packaged-consumer-smoke`.

import { $ } from "bun";
import { cp, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  buildAndPack,
  consumerSrc,
  coreDir,
  reactDir,
  repoRoot,
  writeConsumerPackageJson,
} from "./packaged-consumer-lib";

const PREVIEW_PORT = 4300;
const PREVIEW_URL = `http://localhost:${PREVIEW_PORT}`;

// Poll the preview server until it answers or the deadline passes.
const waitForServer = async (url: string, timeoutMs: number): Promise<boolean> => {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(2_000) });
      if (response.ok || response.status === 404) {
        return true;
      }
    } catch {
      // Server not up yet.
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  return false;
};

let failure: string | null = null;

let corePackDir = "";
let reactPackDir = "";
let consumerDir = "";
let preview: ReturnType<typeof Bun.spawn> | null = null;
try {
  corePackDir = await mkdtemp(path.join(tmpdir(), "folio-core-pack-"));
  reactPackDir = await mkdtemp(path.join(tmpdir(), "folio-react-pack-"));
  consumerDir = await mkdtemp(path.join(tmpdir(), "folio-packaged-consumer-smoke-"));

  const coreTarball = await buildAndPack(coreDir, corePackDir);
  const reactTarball = await buildAndPack(reactDir, reactPackDir);

  // Stage the checked-in consumer OUT of the monorepo so no workspace linkage
  // leaks in. The runtime build needs index.html, the runtime config, the app
  // source, and the served fixture.
  console.log(`→ staging runtime consumer in ${consumerDir}`);
  await cp(path.join(consumerSrc, "src"), path.join(consumerDir, "src"), { recursive: true });
  await cp(path.join(consumerSrc, "public"), path.join(consumerDir, "public"), { recursive: true });
  await cp(path.join(consumerSrc, "index.html"), path.join(consumerDir, "index.html"));
  await cp(
    path.join(consumerSrc, "vite.runtime.config.ts"),
    path.join(consumerDir, "vite.runtime.config.ts"),
  );
  await writeConsumerPackageJson(consumerDir, coreTarball);

  console.log("→ installing tarballs + peers");
  await $`bun add ${reactTarball} ${coreTarball} react@^19 react-dom@^19 use-intl@^4 vite@^8 @vitejs/plugin-react@^6 @types/react@^19 @types/react-dom@^19`
    .cwd(consumerDir)
    .quiet();

  console.log("→ production SPA build over the packed tarballs");
  const build = await $`bun x vite build --config vite.runtime.config.ts`
    .cwd(consumerDir)
    .nothrow();
  if (build.exitCode !== 0) {
    console.error(build.stderr.toString() || build.stdout.toString());
    throw new Error("packaged-consumer-smoke: production vite build FAILED against the tarballs.");
  }

  console.log(`→ serving the built app on ${PREVIEW_URL}`);
  preview = Bun.spawn(
    [
      "bun",
      "x",
      "vite",
      "preview",
      "--config",
      "vite.runtime.config.ts",
      "--port",
      `${PREVIEW_PORT}`,
    ],
    { cwd: consumerDir, stdout: "pipe", stderr: "pipe" },
  );
  if (!(await waitForServer(PREVIEW_URL, 30_000))) {
    throw new Error(`packaged-consumer-smoke: preview server never came up on ${PREVIEW_URL}.`);
  }

  console.log("→ running the runtime smoke (Playwright chromium)");
  const smoke = await $`bun x playwright test --config playwright.smoke.config.ts`
    .cwd(repoRoot)
    .env({ ...process.env, PLAYWRIGHT_BASE_URL: PREVIEW_URL })
    .nothrow();
  if (smoke.exitCode !== 0) {
    failure = "\n✗ packaged-consumer-smoke: the runtime smoke FAILED against the served tarballs.";
  }
} finally {
  preview?.kill();
  for (const dir of [corePackDir, reactPackDir, consumerDir]) {
    if (dir !== "") {
      await rm(dir, { recursive: true, force: true });
    }
  }
}

if (failure !== null) {
  console.error(failure);
  process.exit(1);
}

console.log("\n✓ packaged-consumer-smoke: the served build ran clean (worker + catalog verified).");
