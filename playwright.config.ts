import { defineConfig } from "@playwright/test";
import { PLAYGROUND_HOSTS } from "./tests/parity/playgroundHosts";

const vuePlaygroundPort = new URL(PLAYGROUND_HOSTS.vue).port;

const managedConfig = defineConfig({
  globalSetup: ["./tests/parity/previewSetup.ts", "./tests/parity/playgroundSetup.ts"],
  testDir: "./tests/visual",
  // Bun unit tests share helper directories; browser projects discover only specs.
  testMatch: /\.spec\.ts$/u,
  // A stray test.only must fail CI instead of silently running one test.
  forbidOnly: !!process.env["CI"],
  timeout: 30_000,
  expect: {
    toHaveScreenshot: {
      maxDiffPixelRatio: 0.01, // 1% pixel tolerance (sub-pixel font rounding)
      threshold: 0.2, // per-pixel color sensitivity
      animations: "disabled",
    },
  },
  use: {
    baseURL: PLAYGROUND_HOSTS.react,
    browserName: "chromium",
    viewport: { width: 1280, height: 900 },
    // Consistent rendering across machines
    deviceScaleFactor: 2,
    colorScheme: "light",
  },
  // Projects split the behaviour specs (env-independent: assert content + editor
  // state) from the screenshot baselines (env-specific). CI runs only
  // `--project=interactions` so cross-machine font rendering can't make it
  // flaky; the screenshot baselines stay a local/manual concern.
  //
  // The `parity` project runs the cross-adapter specs in `tests/parity` against
  // both the React (4200) and Vue (4201) playgrounds; `vue` runs only the Vue
  // fork of each parity spec.
  projects: [
    {
      name: "interactions",
      testMatch: /(?:interactions|editing-flows)\.spec\.ts/u,
      testIgnore:
        /(?:(?:canonical-)?browser-input|ai-human-interleaving)-fuzz\.interactions\.spec\.ts/u,
    },
    {
      name: "browser-fuzzer",
      testMatch: /(?:canonical-)?browser-input-fuzz\.interactions\.spec\.ts/u,
    },
    {
      name: "interleaving-fuzzer",
      testMatch: /ai-human-interleaving-fuzz\.interactions\.spec\.ts/u,
    },
    // Measure/paint parity compares two numbers read from the SAME browser in
    // the same layout pass, so unlike the screenshot baselines it cannot go
    // flaky on cross-machine font rendering, and it is safe to gate CI on.
    { name: "measure-parity", testMatch: /measure(?:-backend)?-parity\.spec\.ts/u },
    // Report-only second-engine evidence. These never join a gating lane: the
    // nightly webkit workflow runs `webkit-layout-parity` and records
    // differences. `engine-layout-record` writes the Chromium reference that
    // the same job's WebKit run compares against.
    {
      name: "webkit-layout-parity",
      testMatch: /engine-layout-parity\.spec\.ts/u,
      use: { browserName: "webkit" },
    },
    {
      name: "webkit-measure-backend",
      testMatch: /measure-backend-parity\.spec\.ts/u,
      use: { browserName: "webkit" },
    },
    { name: "engine-layout-record", testMatch: /engine-layout-parity\.spec\.ts/u },
    { name: "rendering", testMatch: /rendering\.spec\.ts/u },
    { name: "performance", testMatch: /editing-performance\.spec\.ts/u },
    {
      name: "parity",
      testDir: "./tests/parity",
      testIgnore: /(?:cross-host|host-api)-flow\.spec\.ts/u,
    },
    {
      name: "parity-fuzzer",
      testDir: "./tests/parity",
      testMatch: /(?:cross-host|host-api)-flow\.spec\.ts/u,
    },
    {
      name: "vue",
      testDir: "./tests/parity",
      grep: /\[vue\]/u,
      testIgnore: /(?:cross-host|host-api)-flow\.spec\.ts/u,
    },
  ],
  // Build fresh workspace sources before serving static previews. Never reuse
  // an existing server: it could be a dev server that reloads on dependency discovery.
  webServer: [
    {
      command:
        "bun scripts/playground-build.ts packages/playground && bun --filter @stll/playground preview",
      stdout: "pipe",
      stderr: "pipe",
      url: PLAYGROUND_HOSTS.react,
      reuseExistingServer: false,
      timeout: 120_000,
    },
    {
      command:
        "bun scripts/playground-build.ts packages/playground-vue && bun --filter @stll/playground-vue preview",
      env: {
        FOLIO_PLAYGROUND_PORT: String(vuePlaygroundPort),
      },
      stdout: "pipe",
      stderr: "pipe",
      url: PLAYGROUND_HOSTS.vue,
      reuseExistingServer: false,
      timeout: 120_000,
    },
  ],
});

export const PLAYGROUND_SERVERS = managedConfig.webServer;

const serverMode = process.env["FOLIO_PLAYGROUND_SERVER_MODE"] ?? "managed-preview";
if (serverMode !== "managed-preview" && serverMode !== "existing-preview") {
  throw new TypeError(`Unknown playground server mode: ${serverMode}`);
}
export const PLAYGROUND_SERVER_MODE = serverMode;

export default defineConfig({
  ...managedConfig,
  webServer: serverMode === "managed-preview" ? PLAYGROUND_SERVERS : [],
});
