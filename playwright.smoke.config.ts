import { defineConfig } from "@playwright/test";

// Playwright config for the packaged-consumer RUNTIME smoke. It reuses the same
// chromium browser Playwright already installs/caches in CI, but is kept
// separate from `playwright.config.ts` (the playground visual/interaction
// suite): this smoke drives the SERVED production build of the packed-tarball
// consumer, whose lifecycle (`vite build` + `vite preview`) is owned by
// `scripts/packaged-consumer-smoke.ts`. That script starts the preview server
// and passes its URL via `PLAYWRIGHT_BASE_URL`, so there is no `webServer` here.
const baseURL = process.env["PLAYWRIGHT_BASE_URL"] ?? "http://localhost:4300";

export default defineConfig({
  testDir: "./test/packaged-consumer",
  testMatch: /smoke\.spec\.ts/u,
  timeout: 60_000,
  forbidOnly: !!process.env["CI"],
  reporter: [["list"]],
  use: {
    baseURL,
    browserName: "chromium",
    viewport: { width: 1280, height: 900 },
  },
});
