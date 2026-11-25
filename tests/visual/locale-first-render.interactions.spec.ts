import { expect, test } from "@playwright/test";

import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";

for (const [adapter, host] of Object.entries(PLAYGROUND_HOSTS)) {
  test(`${adapter} waits for the requested locale before painting editor chrome`, async ({
    page,
  }) => {
    const catalog = /\/cs(?:\.gen)?-[\w-]+\.js$/u;
    const gate = Promise.withResolvers<void>();
    await page.route(catalog, async (route) => {
      await gate.promise;
      await route.continue();
    });

    const requested = page.waitForRequest((request) =>
      catalog.test(request.url()),
    );
    try {
      await page.goto(`${host}/?mountLocale=cs`, {
        waitUntil: "domcontentloaded",
      });
      await requested;
      await expect(page.locator("#app [aria-busy='true']")).toBeVisible();
      await expect(
        page.getByText("Loading document…", { exact: true }),
      ).toHaveCount(0);
      await expect(
        page.getByRole("button", { name: "File", exact: true }),
      ).toHaveCount(0);
    } finally {
      gate.resolve();
    }

    await expect(
      page.getByRole("button", { name: "Soubor", exact: true }),
    ).toBeVisible();
    await expect(
      page.getByRole("button", { name: "File", exact: true }),
    ).toHaveCount(0);
  });
}
