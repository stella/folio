/**
 * Measure/paint parity on a cold font set, with real `unicode-range` subset
 * loading.
 *
 * folio's bundled faces are fontsource packages, split into subsets the
 * browser fetches only when text needs them (Carlito: latin, latin-ext, greek,
 * greek-ext, cyrillic, cyrillic-ext, vietnamese). Every other spec opens a
 * document whose text sits in the Latin subset, or a face installed locally,
 * so none of them saw the first layout measure Czech, Polish, Greek or
 * Cyrillic in a fallback font while the page painted it in the subset that
 * loaded a moment later. Each test here gets a fresh browser context, so
 * nothing is cached or preloaded: the faces load the way they do for a user
 * opening the document.
 */

import { expect, test, type Page } from "@playwright/test";

import { collectLineParity, PER_RUN_TOLERANCE_PX } from "../support/lineParity";

const FIXTURE = "font-subsets.docx";
/** Lines in the fixture at this page width, as a floor against a vacuous pass. */
const MIN_LINES = 18;
/** No layout pass started for this long after the fonts settled: layout is done. */
const LAYOUT_QUIET_MS = 1000;

type LayoutLogEntry = {
  reason: string;
  /** Carlito regular's subsets and their load status when the pass started. */
  carlito: Record<string, string>;
};

declare global {
  var __fontSubsetLayoutLog: LayoutLogEntry[] | undefined;
  var __fontSubsetStatus: (() => Record<string, string>) | undefined;
}

/** Record, for every layout pass, why it ran and which subsets had loaded. */
async function installLayoutLog(page: Page): Promise<void> {
  await page.addInitScript(() => {
    // fontsource's subsets, named by where their `unicode-range` starts.
    const SUBSET_BY_RANGE_START: [string, string][] = [
      ["U+460", "cyrillic-ext"],
      ["U+301", "cyrillic"],
      ["U+1F00", "greek-ext"],
      ["U+370", "greek"],
      ["U+102", "vietnamese"],
      ["U+100", "latin-ext"],
      ["U+0-FF", "latin"],
    ];
    const subsetOf = (range: string): string =>
      SUBSET_BY_RANGE_START.find(([start]) => range.startsWith(start))?.[1] ?? range;
    globalThis.__fontSubsetStatus = () => {
      const status: Record<string, string> = {};
      for (const face of document.fonts) {
        if (
          face.family.replaceAll('"', "") === "Carlito" &&
          face.weight === "400" &&
          face.style === "normal"
        ) {
          status[subsetOf(face.unicodeRange)] = face.status;
        }
      }
      return status;
    };
    globalThis.__fontSubsetLayoutLog = [];
    globalThis.__folioLayoutInstrumentation = {
      onLayoutStart: ({ reason }) => {
        globalThis.__fontSubsetLayoutLog?.push({
          reason,
          carlito: globalThis.__fontSubsetStatus?.() ?? {},
        });
      },
    };
  });
}

/** Resolve once the fonts are settled and no layout pass has started for a while. */
async function settleLayout(page: Page): Promise<LayoutLogEntry[]> {
  const passesOnceFontsSettle = () =>
    page.evaluate(async () => {
      await document.fonts.ready;
      return globalThis.__fontSubsetLayoutLog?.length ?? 0;
    });
  let passes = await passesOnceFontsSettle();
  let previous: number;
  do {
    previous = passes;
    await page.waitForTimeout(LAYOUT_QUIET_MS);
    passes = await passesOnceFontsSettle();
  } while (passes !== previous);
  return page.evaluate(() => globalThis.__fontSubsetLayoutLog ?? []);
}

async function expectMeasuredAsPainted(page: Page): Promise<void> {
  const lines = await collectLineParity(page);
  expect(lines.length).toBeGreaterThanOrEqual(MIN_LINES);
  expect(
    lines
      .filter((line) => Math.abs(line.delta) > PER_RUN_TOLERANCE_PX * Math.max(1, line.runCount))
      .map(
        (line) =>
          `line ${line.index} "${line.text}": measured ${line.measured.toFixed(2)}px, painted ${line.painted.toFixed(2)}px`,
      ),
  ).toEqual([]);
}

test.describe("bundled font subsets on a cold font set", () => {
  test.beforeEach(async ({ page }) => {
    await installLayoutLog(page);
    await page.goto(`/?file=${FIXTURE}`);
    await page.waitForSelector(".layout-page .layout-line", { timeout: 30_000 });
  });

  test("the first layout waits for every subset the text needs, and nothing relays out", async ({
    page,
  }) => {
    const log = await settleLayout(page);

    const initial = log.find((entry) => entry.reason === "initial");
    expect(initial?.carlito).toMatchObject({
      latin: "loaded",
      "latin-ext": "loaded",
      greek: "loaded",
      cyrillic: "loaded",
    });
    // The fixture has no Vietnamese letter; waiting for text, not for every
    // face, leaves that subset alone.
    expect(initial?.carlito["vietnamese"]).toBe("unloaded");
    expect(log.filter((entry) => entry.reason === "font-ready")).toEqual([]);
    await expectMeasuredAsPainted(page);
  });

  test("a subset first needed after the first layout re-measures the lines that use it", async ({
    page,
  }) => {
    await settleLayout(page);
    expect(await page.evaluate(() => globalThis.__fontSubsetStatus?.()["vietnamese"])).toBe(
      "unloaded",
    );

    await page.locator(".layout-paragraph").first().click();
    await page.keyboard.press("End");
    // Synthetic Vietnamese: the letters live only in Carlito's vietnamese
    // subset. One insertion, so one layout pass measures it before the subset
    // loads and no later keystroke re-measures it: only the font-load
    // follow-up can.
    await page.keyboard.insertText(
      " Hợp đồng được ký kết giữa các bên và có hiệu lực kể từ ngày ký.",
    );
    await settleLayout(page);

    expect(await page.evaluate(() => globalThis.__fontSubsetStatus?.()["vietnamese"])).toBe(
      "loaded",
    );
    await expectMeasuredAsPainted(page);
  });
});
