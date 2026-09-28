import { expect, test } from "@playwright/test";

import { collectLineParity, PER_RUN_TOLERANCE_PX } from "../support/lineParity";

const FIXTURE = "mixed-script-font-measure.docx";
const EDITOR_HOSTS = [
  { name: "React", url: "http://localhost:4200" },
  { name: "Vue", url: "http://localhost:4201" },
] as const;

for (const host of EDITOR_HOSTS) {
  test(`${host.name} measures mixed-script text with its loaded fonts`, async ({ page }) => {
    const failedFontRequests: string[] = [];
    page.on("response", (response) => {
      if (/\.woff2?(?:\?|$)/u.test(response.url()) && response.status() >= 400) {
        failedFontRequests.push(`${response.status()} ${response.url()}`);
      }
    });
    await page.goto(`${host.url}/?file=${FIXTURE}`);
    await page.waitForSelector(".layout-page .layout-line", { timeout: 30_000 });
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
