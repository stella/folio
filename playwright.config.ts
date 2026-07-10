import { defineConfig } from "@playwright/test";

const reactPlaygroundPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;

export default defineConfig({
  testDir: "./tests/visual",
  timeout: 30_000,
  expect: {
    toHaveScreenshot: {
      maxDiffPixelRatio: 0.01, // 1% pixel tolerance (sub-pixel font rounding)
      threshold: 0.2, // per-pixel color sensitivity
      animations: "disabled",
    },
  },
  use: {
    baseURL: `http://localhost:${reactPlaygroundPort}`,
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
    { name: "interactions", testMatch: /(?:interactions|editing-flows)\.spec\.ts/u },
    { name: "rendering", testMatch: /rendering\.spec\.ts/u },
    { name: "performance", testMatch: /editing-performance\.spec\.ts/u },
    { name: "parity", testDir: "./tests/parity" },
    { name: "vue", testDir: "./tests/parity", grep: /\[vue\]/u },
  ],
  // Start both playground dev servers automatically (reused if already running).
  // The React server backs the visual/interaction suites; the Vue server backs
  // the parity project. Both boot for any run — `reuseExistingServer` keeps a
  // manually-started dev server in place.
  webServer: [
    {
      command: "bun --filter @stll/playground dev",
      url: `http://localhost:${reactPlaygroundPort}`,
      reuseExistingServer: true,
      timeout: 120_000,
    },
    {
      command: "bun --filter @stll/playground-vue dev",
      env: {
        FOLIO_PLAYGROUND_PORT: "4201",
      },
      url: "http://localhost:4201",
      reuseExistingServer: true,
      timeout: 120_000,
    },
  ],
});
