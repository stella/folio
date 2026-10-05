import { expect, test } from "@playwright/test";
import fs from "node:fs/promises";
import config from "../../playwright.config";

const servers = config.webServer;
if (!Array.isArray(servers)) throw new TypeError("Expected both playground servers.");

for (const server of servers) {
  test(`built playground and fixture transport: ${server.url}`, async ({ page, request }) => {
    if (!server.url) throw new TypeError("Missing playground URL.");
    const response = await request.get(server.url);
    expect(response.ok()).toBe(true);
    const html = await response.text();
    expect(html).not.toContain("/@vite/client");
    expect(html).toMatch(/src="\/assets\/[^" ]+\.js"/u);

    const fixture = await request.get(`${server.url}/fixtures/sample.docx`);
    expect(fixture.status()).toBe(200);
    expect(fixture.headers()["content-type"]).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    expect(await fixture.body()).toEqual(
      await fs.readFile(new URL("./fixtures/sample.docx", import.meta.url)),
    );
    for (const path of ["%", "a%2Fb.docx", "a%5Cb.docx", "..%2Fx.docx"]) {
      expect((await request.get(`${server.url}/fixtures/${path}`)).status()).toBe(400);
    }
    expect((await request.get(`${server.url}/fixtures/missing.docx`)).status()).toBe(404);

    await page.goto(`${server.url}/?file=sample.docx`);
    await expect(page.locator(".paged-editor__pages").first()).toBeVisible();
    expect(await page.locator('script[src^="/assets/"]').count()).toBeGreaterThan(0);
  });
}
