// Scratch measurement only: engine timings against a running preview (FOLIO_ZZ_BASE).
import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/visual",
  testMatch: /zz-flip-measure\.spec\.ts$/u,
  timeout: 600_000,
  workers: 1,
  fullyParallel: false,
  reporter: "line",
  use: {
    baseURL: process.env["FOLIO_ZZ_BASE"] ?? "http://localhost:4200",
    browserName: "chromium",
    viewport: { width: 1280, height: 900 },
    deviceScaleFactor: 1,
    colorScheme: "light",
  },
});
