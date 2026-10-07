import { expect } from "@playwright/test";
import { assertCanonicalInputTimersSettled, test } from "./canonicalTimerProbe";

// Use a real script resource so the browser emits the same owner-frame shape
// as canonical input code, rather than an anonymous page.evaluate frame.
const bootstrapURL = "http://localhost:4200/timer-probe";
const ownerURL = "http://localhost:4200/packages/core/src/controller/canonicalInput.ts";

test("the first assertion sees timers scheduled during document startup and reload", async ({
  page,
}) => {
  await page.route(bootstrapURL, (route) =>
    route.fulfill({
      contentType: "text/html",
      body: `<script src="${ownerURL}"></script>`,
    }),
  );
  await page.route(ownerURL, (route) =>
    route.fulfill({
      contentType: "application/javascript",
      body: "globalThis.__earlyCanonicalTimer = setTimeout(() => {}, 60000);",
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
    await assertCanonicalInputTimersSettled(page);
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
