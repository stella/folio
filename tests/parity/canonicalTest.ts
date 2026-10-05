import { test as base, type Frame } from "@playwright/test";

/** Preserve navigation evidence when an evaluation loses its document context. */
export const test = base.extend<{ canonicalNavigationEvidence: void }>({
  canonicalNavigationEvidence: [
    async ({ page }, use, testInfo) => {
      const started = Date.now();
      const navigations: { elapsedMs: number; url: string; frame: "main" | "child" }[] = [];
      const record = (frame: Frame) => {
        navigations.push({
          elapsedMs: Date.now() - started,
          url: frame.url(),
          frame: frame === page.mainFrame() ? "main" : "child",
        });
      };
      page.on("framenavigated", record);
      try {
        await use();
        if (testInfo.status !== testInfo.expectedStatus) {
          const body = JSON.stringify(navigations, null, 2);
          console.log(`Canonical browser navigation evidence (${testInfo.title}):\n${body}`);
          await testInfo.attach("canonical-navigations", { body, contentType: "application/json" });
        }
      } finally {
        page.off("framenavigated", record);
      }
    },
    { auto: true },
  ],
});
