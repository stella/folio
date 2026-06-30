import { defineConfig } from "@playwright/test";

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
    baseURL: "http://localhost:4200",
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
  projects: [
    { name: "interactions", testMatch: /(?:interactions|editing-flows)\.spec\.ts/u },
    { name: "rendering", testMatch: /rendering\.spec\.ts/u },
    { name: "performance", testMatch: /editing-performance\.spec\.ts/u },
  ],
  // Start the playground dev server automatically (reused if already running).
  webServer: {
    command: "bun --filter @stll/playground dev",
    url: "http://localhost:4200",
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
