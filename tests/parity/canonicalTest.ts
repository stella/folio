import { test as base, type Frame } from "@playwright/test";
import { readFile } from "node:fs/promises";
import { observeCanonicalPageNavigation } from "../visual/canonicalPageNavigation";

/** Preserve navigation evidence when an evaluation loses its document context. */
export const test = base.extend<{ canonicalNavigationEvidence: void }>({
  canonicalNavigationEvidence: [
    async ({ page }, use, testInfo) => {
      observeCanonicalPageNavigation(page);
      const cdp = await page.context().newCDPSession(page);
      const contexts: { elapsedMs: number; method: string; payload: unknown }[] = [];
      const started = Date.now();
      const capture = (method: string, payload: unknown) =>
        contexts.push({ elapsedMs: Date.now() - started, method, payload });
      cdp.on("Runtime.executionContextCreated", (payload) =>
        capture("Runtime.executionContextCreated", payload),
      );
      cdp.on("Runtime.executionContextDestroyed", (payload) =>
        capture("Runtime.executionContextDestroyed", payload),
      );
      cdp.on("Runtime.executionContextsCleared", (payload) =>
        capture("Runtime.executionContextsCleared", payload),
      );
      cdp.on("Inspector.targetCrashed", (payload) => capture("Inspector.targetCrashed", payload));
      await cdp.send("Runtime.enable");
      const navigations: { elapsedMs: number; url: string; frame: "main" | "child" }[] = [];
      const record = (frame: Frame) => {
        navigations.push({
          elapsedMs: Date.now() - started,
          url: frame.url(),
          frame: frame === page.mainFrame() ? "main" : "child",
        });
      };
      page.on("framenavigated", record);
      let detachFailure: { error: unknown } | null = null;
      try {
        await use();
        if (testInfo.status !== testInfo.expectedStatus) {
          const body = JSON.stringify(navigations, null, 2);
          console.log(`Canonical browser navigation evidence (${testInfo.title}):\n${body}`);
          await testInfo.attach("canonical-navigations", { body, contentType: "application/json" });
          await testInfo.attach("canonical-execution-contexts", {
            body: JSON.stringify(contexts, null, 2),
            contentType: "application/json",
          });
          const protocolPath = process.env["FOLIO_CANONICAL_PROTOCOL_LOG"];
          if (protocolPath)
            await testInfo.attach("canonical-protocol", {
              body: await readFile(protocolPath),
              contentType: "application/x-ndjson",
            });
        }
      } finally {
        page.off("framenavigated", record);
        if (!page.isClosed()) {
          try {
            await cdp.detach();
          } catch (error) {
            detachFailure = { error };
            await testInfo.attach("canonical-protocol-detach-error", {
              body: String(error),
              contentType: "text/plain",
            });
          }
        }
      }
      if (detachFailure && testInfo.status === testInfo.expectedStatus) throw detachFailure.error;
    },
    { auto: true },
  ],
});
