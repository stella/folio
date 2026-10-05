import type {} from "./browserTestBridge";
import { expect, test } from "@playwright/test";
import fs from "node:fs/promises";
import config, { PLAYGROUND_SERVERS, PLAYGROUND_SERVER_MODE } from "../../playwright.config";
import reactPlayground from "../../packages/playground/package.json" with { type: "json" };
import vuePlayground from "../../packages/playground-vue/package.json" with { type: "json" };

const servers = PLAYGROUND_SERVERS;
if (!Array.isArray(servers)) throw new TypeError("Expected both playground servers.");

test("browser harness owns a fresh preview for both playground packages", () => {
  expect(config.webServer).toEqual(PLAYGROUND_SERVER_MODE === "existing-preview" ? [] : servers);
  const packages = [reactPlayground, vuePlayground];
  expect(servers).toHaveLength(packages.length);
  expect(servers.map(({ command }) => command).sort()).toEqual(
    packages
      .map(
        ({ name }) =>
          `bun scripts/playground-build.ts packages/${name.replace("@stll/", "")} && bun --filter ${name} preview`,
      )
      .sort(),
  );
  for (const server of servers) expect(server.reuseExistingServer).toBe(false);
});

for (const server of servers) {
  test(`built playground and fixture transport: ${server.url}`, async ({ page, request }) => {
    if (!server.url) throw new TypeError("Missing playground URL.");
    const response = await request.get(server.url);
    expect(response.ok()).toBe(true);
    const html = await response.text();
    expect(html).not.toMatch(/[/]@vite[/]client/u);
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
    await expect(page.locator(".paged-editor__pages").first()).toContainText(/\S/u);
    if (server.url === servers.at(0)?.url) {
      expect(
        await page.evaluate(() => {
          const bridge = globalThis.__folioBrowserTestBridge;
          if (!bridge) throw new Error("bundled browser test bridge unavailable");
          bridge.installInterleavingBridge();
          return Object.keys(bridge).sort();
        }),
      ).toEqual([
        "installInterleavingBridge",
        "resolvePaintedDragTarget",
        "resolvePaintedTableTarget",
      ]);
    }
    expect(await page.locator('script[src^="/assets/"]').count()).toBeGreaterThan(0);
  });
}
