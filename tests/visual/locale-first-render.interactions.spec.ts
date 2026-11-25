import { expect, test } from "@playwright/test";

import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";

const czechCatalog = /\/cs(?:\.gen)?-[\w-]+\.js$/u;

for (const [adapter, host] of Object.entries(PLAYGROUND_HOSTS)) {
  test(`${adapter} waits for the requested locale before painting editor chrome`, async ({
    page,
  }) => {
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(error.message));
    const gate = Promise.withResolvers<void>();
    await page.route(czechCatalog, async (route) => {
      await gate.promise;
      await route.continue();
    });

    const requested = page.waitForRequest((request) =>
      czechCatalog.test(request.url()),
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
        page.getByRole("button", { name: "Bold", exact: true }),
      ).toHaveCount(0);
    } finally {
      gate.resolve();
    }

    await expect(
      page.getByRole("button", { name: "Tučné", exact: true }),
    ).toBeVisible();
    await expect(
      page.getByRole("button", { name: "Bold", exact: true }),
    ).toHaveCount(0);
    expect(errors).toEqual([]);
  });
}

test("a directly mounted Vue editor shows failure when its locale cannot load", async ({
  page,
}) => {
  await page.route(czechCatalog, (route) => route.abort());
  await page.goto(`${PLAYGROUND_HOSTS.vue}/?locale=cs&file=sample.docx`);
  const editor = page.locator(".docx-editor-vue");
  await expect(editor.getByRole("alert")).toBeVisible();
  await expect(editor.locator("[aria-busy='true']")).toHaveCount(0);
  await expect(
    editor.getByRole("button", { name: "Bold", exact: true }),
  ).toHaveCount(0);
});
