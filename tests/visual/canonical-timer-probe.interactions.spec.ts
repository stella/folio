import { expect } from "@playwright/test";
import { readFileSync } from "node:fs";
import ts from "typescript";
import { assertCanonicalInputTimersSettled, test } from "./canonicalTimerProbe";

// Serve the actual producer as a browser module with a bundle-shaped URL.
// Ownership must survive the absence of source paths in browser stacks.
const bootstrapURL = "http://localhost:4200/timer-probe";
const bootstrapScriptURL = "http://localhost:4200/assets/bootstrap-123.js";
const ownerURL = "http://localhost:4200/assets/input-timer-456.js";
const ownerModule = ts.transpileModule(
  readFileSync(
    new URL("../../packages/core/src/controller/canonicalInputTimer.ts", import.meta.url),
    "utf8",
  ),
  { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022 } },
).outputText;

test("the first assertion sees timers scheduled during document startup and reload", async ({
  page,
}) => {
  await page.route(bootstrapURL, (route) =>
    route.fulfill({
      contentType: "text/html",
      body: `<script type="module" src="${bootstrapScriptURL}"></script>`,
    }),
  );
  await page.route(ownerURL, (route) =>
    route.fulfill({
      contentType: "application/javascript",
      body: ownerModule,
    }),
  );
  await page.route(bootstrapScriptURL, (route) =>
    route.fulfill({
      contentType: "application/javascript",
      body: `import { setCanonicalInputTimer } from "${ownerURL}";
      const callback = () => {};
      globalThis.__earlyCanonicalTimer = setCanonicalInputTimer(callback, 60000);
      globalThis.__unownedTimer = setTimeout(callback, 60000);`,
    }),
  );
  await page.goto(bootstrapURL);
  for (let navigation = 0; navigation < 2; navigation++) {
    await expect(assertCanonicalInputTimersSettled(page)).rejects.toThrow(
      "case must start with no canonical input timer",
    );
    await page.evaluate(() => {
      const timer = Reflect.get(globalThis, "__earlyCanonicalTimer");
      if (typeof timer !== "number") throw new TypeError("Startup timer unavailable");
      window.clearTimeout(timer);
    });
    // The unrelated timer remains pending and must not be attributed to input.
    await assertCanonicalInputTimersSettled(page);
    await page.evaluate(() => {
      const timer = Reflect.get(globalThis, "__unownedTimer");
      if (typeof timer !== "number") throw new TypeError("Unowned timer unavailable");
      window.clearTimeout(timer);
    });
    if (navigation === 0) await page.reload();
  }
});

test("an uninstrumented document fails instead of reporting no pending timers", async ({
  browser,
}) => {
  const page = await browser.newPage();
  try {
    await page.goto("data:text/html,<p>uninstrumented</p>");
    await expect(assertCanonicalInputTimersSettled(page)).rejects.toThrow(
      "Canonical timer instrumentation must be installed before navigation",
    );
  } finally {
    await page.close();
  }
});
