#!/usr/bin/env bun
// The smoke test in a real VS Code: download VS Code (cached in .vscode-test/),
// load the built extension from this folder, open a copy of a fixture .docx
// in a fresh workspace and profile, and run `suite.ts` inside the editor.
//
// Run `bun run build` first. Without a display (Linux CI), run it under
// `xvfb-run -a`. FOLIO_SMOKE_VSCODE_VERSION picks the VS Code version
// (default: the latest stable).

import { runTests } from "@vscode/test-electron";
import { build } from "esbuild";
import { copyFile, mkdtemp, realpath, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const extensionRoot = path.resolve(import.meta.dir, "..", "..");
const repoRoot = path.resolve(extensionRoot, "..", "..");
const suiteOut = path.join(extensionRoot, "out", "smoke", "suite.js");
const FIXTURE = path.join(repoRoot, "packages", "playground", "public", "folio-showcase.docx");

await stat(path.join(extensionRoot, "dist", "extension.js")).catch(() => {
  throw new Error("Build the extension first: bun run build");
});

await build({
  entryPoints: [path.join(import.meta.dir, "suite.ts")],
  outfile: suiteOut,
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node22",
  external: ["vscode"],
  logLevel: "warning",
});

const workspace = await realpath(await mkdtemp(path.join(tmpdir(), "folio-smoke-workspace-")));
const userData = await realpath(await mkdtemp(path.join(tmpdir(), "folio-smoke-profile-")));
const documentPath = path.join(workspace, "Smoke test.docx");
await copyFile(FIXTURE, documentPath);

// A VS Code that hangs (a dialog nobody answers) fails the run instead of the job.
const TIMEOUT_MS = 10 * 60 * 1000;
const watchdog = setTimeout(() => {
  console.error(`The smoke test did not finish within ${String(TIMEOUT_MS / 60_000)} minutes.`);
  process.exit(1);
}, TIMEOUT_MS);

try {
  await runTests({
    version: process.env["FOLIO_SMOKE_VSCODE_VERSION"] ?? "stable",
    cachePath: path.join(extensionRoot, ".vscode-test"),
    extensionDevelopmentPath: extensionRoot,
    extensionTestsPath: suiteOut,
    extensionTestsEnv: { FOLIO_VSCODE_TEST: "1", FOLIO_SMOKE_DOCUMENT: documentPath },
    launchArgs: [
      workspace,
      "--user-data-dir",
      userData,
      "--disable-extensions",
      "--disable-workspace-trust",
      "--skip-welcome",
      "--skip-release-notes",
      // CI runners cannot give Chromium its sandbox.
      ...(process.platform === "linux" ? ["--no-sandbox", "--disable-gpu"] : []),
    ],
  });
} finally {
  clearTimeout(watchdog);
  if (process.env["FOLIO_SMOKE_KEEP"] === "1") {
    // The profile's logs/ holds the extension host's and the webviews' logs.
    console.log(`Kept ${workspace} and ${userData}`);
  } else {
    await rm(workspace, { recursive: true, force: true });
    await rm(userData, { recursive: true, force: true });
  }
}
