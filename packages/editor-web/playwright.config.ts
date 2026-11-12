import { defineConfig } from "@playwright/test";

import { withBoundedProtocolCapture } from "../../tests/support/playwrightProtocolCapture";

// Runs the built VS Code bundle (`bun run build` first; `test:e2e` does both)
// in a page carrying the webview's Content Security Policy. No server: the
// spec answers every request itself.
export default withBoundedProtocolCapture(
  defineConfig({
    testDir: "./tests",
    timeout: 90_000,
    use: {
      browserName: "chromium",
      viewport: { width: 1280, height: 900 },
    },
  }),
);
