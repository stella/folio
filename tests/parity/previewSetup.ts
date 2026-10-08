import { expect, request } from "@playwright/test";
import fs from "node:fs/promises";
import { PLAYGROUND_SERVERS } from "../../playwright.config";

/** Validate managed and explicitly supplied hosts before any test owns a page. */
export default async () => {
  if (!Array.isArray(PLAYGROUND_SERVERS)) throw new TypeError("Missing playground servers.");
  const client = await request.newContext();
  try {
    const expectedFixture = await fs.readFile(
      new URL("../visual/fixtures/sample.docx", import.meta.url),
    );
    for (const { url } of PLAYGROUND_SERVERS) {
      if (!url) throw new TypeError("Missing playground URL.");
      const index = await client.get(url);
      expect(index.ok(), `Preview unavailable: ${url}`).toBe(true);
      const html = await index.text();
      expect(html, `Development server supplied instead of built preview: ${url}`).not.toContain(
        "/@vite/client",
      );
      expect(html, `Missing built assets: ${url}`).toMatch(/src="\/assets\/[^" ]+\.js"/u);
      const fixture = await client.get(`${url}/fixtures/sample.docx`);
      expect(fixture.status(), `Fixture middleware unavailable: ${url}`).toBe(200);
      expect(await fixture.body(), `Fixture bytes differ: ${url}`).toEqual(expectedFixture);
    }
  } finally {
    await client.dispose();
  }
};
