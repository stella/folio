import { defineConfig } from "@playwright/test";

import { withBoundedProtocolCapture } from "../../../tests/support/playwrightProtocolCapture";

// A browserless run whose only test floods Playwright's protocol logger.
export default withBoundedProtocolCapture(
  defineConfig({
    testDir: ".",
    testMatch: /protocol-flood\.pw\.ts/u,
    reporter: [["json"]],
    workers: 1,
  }),
);
