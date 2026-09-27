import { expect, test } from "@playwright/test";

import { collectLineParity, PER_RUN_TOLERANCE_PX } from "../support/lineParity";

const FIXTURE = "mixed-script-font-sample.docx";
const EDITOR_HOSTS = [
  { name: "React", url: "http://localhost:4200" },
  { name: "Vue", url: "http://localhost:4201" },
] as const;

for (const host of EDITOR_HOSTS) {
  test(`${host.name} measures mixed-script text with its loaded fonts`, async ({ page }) => {
    const failedFontRequests: string[] = [];
    const failedResponses: string[] = [];
    const failedRequests: string[] = [];
    const pageErrors: string[] = [];
    const consoleErrors: string[] = [];
    page.on("response", (response) => {
      if (response.status() >= 400) {
        const path = new URL(response.url()).pathname;
        failedResponses.push(`${response.status()} ${path}`);
        if (/\.woff2?$/u.test(path)) {
          failedFontRequests.push(`${response.status()} ${path}`);
        }
      }
    });
    page.on("requestfailed", (request) => {
      const path = new URL(request.url()).pathname;
      failedRequests.push(`${path}: ${request.failure()?.errorText ?? "unknown failure"}`);
    });
    page.on("pageerror", (error) => pageErrors.push(error.message.slice(0, 300)));
    page.on("console", (message) => {
      if (message.type() === "error") {
        consoleErrors.push(message.text().slice(0, 300));
      }
    });
    await page.goto(`${host.url}/?file=${FIXTURE}`);
    try {
      await page.waitForSelector(".layout-page .layout-line", { timeout: 30_000 });
    } catch (error) {
      const domSummary = await page.evaluate(() => ({
        bodyText: document.body.innerText.slice(0, 600),
        lineCount: document.querySelectorAll(".layout-page .layout-line").length,
        pageCount: document.querySelectorAll(".layout-page").length,
        statusText: [
          ...document.querySelectorAll(
            ".pg-vue-status, .docx-editor-vue__error, .docx-editor-vue__loading",
          ),
        ].map((element) => element.textContent?.slice(0, 200) ?? ""),
      }));
      throw new Error(
        `${host.name} layout did not render: ${JSON.stringify({
          domSummary,
          failedResponses: failedResponses.slice(-8),
          failedRequests: failedRequests.slice(-8),
          pageErrors: pageErrors.slice(-8),
          consoleErrors: consoleErrors.slice(-8),
        })}`,
        { cause: error },
      );
    }
    await page.evaluate(() => document.fonts.ready);
    expect(failedFontRequests).toEqual([]);
    const loadedFamilies = await page.evaluate(() =>
      [...document.fonts].filter((face) => face.status === "loaded").map((face) => face.family),
    );
    expect(loadedFamilies.some((family) => family.includes("Arimo Embedded"))).toBe(true);
    expect(loadedFamilies).toContain("Noto Sans Arabic");

    const lines = await collectLineParity(page);
    expect(lines.length).toBeGreaterThan(30);
    expect(lines.some((line) => /[\u0590-\u05ff]/u.test(line.text))).toBe(true);
    expect(lines.some((line) => /[\u0600-\u06ff]/u.test(line.text))).toBe(true);
    expect(lines.some((line) => /[\u3400-\u9fff]/u.test(line.text))).toBe(true);
    expect(
      lines
        .filter((line) => Math.abs(line.delta) > PER_RUN_TOLERANCE_PX * line.runCount)
        .map(
          (line) =>
            `line ${line.index} "${line.text}": measured ${line.measured.toFixed(2)}px, painted ${line.painted.toFixed(2)}px`,
        ),
    ).toEqual([]);
  });
}
