#!/usr/bin/env bun
// Runtime half of the packaged-consumer gate: build + pack the workspace
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
import { cp, mkdtemp, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { installCleanRoomDependencies } from "./clean-room-install";
import { assertHostPeerResolution } from "./host-peer-policy";
import {
  REACT_PEER_FLOOR,
  buildAndPack,
  consumerSrc,
  coreDir,
  docxCoreDir,
  isDocxKernelWasmAsset,
  hostPeerInstallArgs,
  reactPeerInstallArgs,
  reactDir,
  repoRoot,
  writeConsumerPackageJson,
} from "./packaged-consumer-lib";

const previewPort = 4300 + Number(REACT_PEER_FLOOR.major);
const previewUrl = `http://localhost:${previewPort}`;
const runtimeDistDirectory = "dist-runtime";

// Poll the preview server until it answers or the deadline passes.
const waitForServer = async (url: string, timeoutMs: number): Promise<boolean> => {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(2_000) });
      // Drain the body so each poll releases its connection instead of
      // leaving a socket tied to an unconsumed response.
      await response.arrayBuffer();
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

// Print a captured preview-server log with a header, if it has content.
const dumpLog = async (label: string, logPath: string): Promise<void> => {
  if (logPath === "") {
    return;
  }
  const file = Bun.file(logPath);
  if (!(await file.exists())) {
    return;
  }
  const content = await file.text();
  if (content.trim() !== "") {
    console.error(`\n--- ${label} ---\n${content}`);
  }
};

let failure: string | null = null;

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

  const runSmoke = async (reactVersion: string): Promise<void> => {
    let succeeded = false;
    let previewStdoutPath = "";
    let previewStderrPath = "";
    let preview: ReturnType<typeof Bun.spawn> | null = null;
    const consumerDir = await mkdtemp(
      path.join(tmpdir(), `folio-packaged-consumer-smoke-r${reactVersion}-`),
    );
    consumerDirs.push(consumerDir);

    try {
      // Stage the checked-in consumer OUT of the monorepo so no workspace linkage
      // leaks in. The runtime build needs index.html, the runtime config, the app
      // source, and the served fixture.
      console.log(`→ staging React ${reactVersion} runtime consumer in ${consumerDir}`);
      await cp(path.join(consumerSrc, "src"), path.join(consumerDir, "src"), { recursive: true });
      await cp(path.join(consumerSrc, "public"), path.join(consumerDir, "public"), {
        recursive: true,
      });
      await cp(path.join(consumerSrc, "index.html"), path.join(consumerDir, "index.html"));
      await cp(
        path.join(consumerSrc, "vite.runtime.config.ts"),
        path.join(consumerDir, "vite.runtime.config.ts"),
      );
      await writeConsumerPackageJson({ consumerDir, coreTarball, docxCoreTarball });

      console.log(`→ installing tarballs + React ${reactVersion} peers`);
      await installCleanRoomDependencies({
        consumerDir,
        dependencies: [
          reactTarball,
          coreTarball,
          docxCoreTarball,
          ...reactPeerInstallArgs(),
          ...hostPeerInstallArgs(),
          "use-intl@^4",
          "vite@^8",
          "@vitejs/plugin-react@^6",
        ],
      });

      assertHostPeerResolution(consumerDir);
      console.log(`→ production SPA build over the packed tarballs (React ${reactVersion})`);
      const build = await $`bun x vite build --config vite.runtime.config.ts`
        .cwd(consumerDir)
        .nothrow();
      if (build.exitCode !== 0) {
        console.error(build.stderr.toString() || build.stdout.toString());
        throw new Error(
          `packaged-consumer-smoke: React ${reactVersion} production vite build FAILED against the tarballs.`,
        );
      }
      const distFiles = await readdir(path.join(consumerDir, runtimeDistDirectory), {
        recursive: true,
      });
      if (!distFiles.some(isDocxKernelWasmAsset)) {
        throw new Error(
          `packaged-consumer-smoke: React ${reactVersion} production app emitted no lazy DOCX kernel WebAssembly asset.`,
        );
      }

      console.log(`→ serving the React ${reactVersion} built app on ${previewUrl}`);
      // Redirect the server's stdio to files rather than piping: an unread pipe
      // blocks the child once the OS buffer fills, and the files double as CI
      // debugging output; they are dumped below whenever the smoke fails.
      previewStdoutPath = path.join(consumerDir, "vite-preview.stdout.log");
      previewStderrPath = path.join(consumerDir, "vite-preview.stderr.log");
      preview = Bun.spawn(
        [
          "bun",
          "x",
          "vite",
          "preview",
          "--config",
          "vite.runtime.config.ts",
          "--port",
          `${previewPort}`,
        ],
        {
          cwd: consumerDir,
          stdout: Bun.file(previewStdoutPath),
          stderr: Bun.file(previewStderrPath),
        },
      );
      if (!(await waitForServer(previewUrl, 30_000))) {
        throw new Error(
          `packaged-consumer-smoke: React ${reactVersion} preview server never came up on ${previewUrl}.`,
        );
      }

      console.log(`→ running the React ${reactVersion} runtime smoke (Playwright chromium)`);
      const smoke = await $`bun x playwright test --config playwright.smoke.config.ts`
        .cwd(repoRoot)
        .env({ ...process.env, PLAYWRIGHT_BASE_URL: previewUrl })
        .nothrow();
      if (smoke.exitCode === 0) {
        succeeded = true;
        return;
      }
      failure = `\n✗ packaged-consumer-smoke: the React ${reactVersion} runtime smoke FAILED against the served tarballs.`;
    } finally {
      preview?.kill();
      // On any failure (assertion or thrown step), surface the preview server's
      // captured output for CI debugging before the staging dir is removed.
      if (!succeeded) {
        await dumpLog("vite preview stdout", previewStdoutPath);
        await dumpLog("vite preview stderr", previewStderrPath);
      }
    }
  };

  await runSmoke(REACT_PEER_FLOOR.version);
} finally {
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
  `\n✓ packaged-consumer-smoke: React ${REACT_PEER_FLOOR.version} served build ran clean (worker + catalog verified).`,
);
