/**
 * Layout under a second browser engine: record where it differs, do not gate.
 *
 * Folio breaks lines and paginates itself, but glyph advances come from the
 * browser's canvas `measureText`. A different engine can therefore break a line
 * in a different place. This spec records, per fixture document, the page
 * count and for every painted line its first and last word and its measured
 * width, then compares that against the record taken under Chromium and
 * committed in `engine-layout-parity-baseline/`.
 *
 * It is report-only. A difference never fails a test: each document writes a
 * JSON file with its structured differences into `FOLIO_ENGINE_PARITY_OUT`
 * (default `engine-parity-out`), and a test fails only when the harness cannot
 * load or read the document.
 *
 * Set `FOLIO_ENGINE_PARITY_RECORD=1` under Chromium to (re)write the baseline.
 *
 * It also records whether `ctx.fontKerning = "none"` changes `measureText`
 * under the running engine, since measurement and paint then disagree on
 * kerning, and which font stack each document actually resolved.
 */

import fs from "node:fs";
import path from "node:path";

import { test, expect, type Page } from "@playwright/test";

import {
  diffRecords,
  type DocumentRecord,
  type FontResolution,
  type PageRecord,
} from "./engineLayoutDiff";

const BASELINE_DIR = path.resolve("tests", "visual", "engine-layout-parity-baseline");
const OUT_DIR = path.resolve(process.env["FOLIO_ENGINE_PARITY_OUT"] ?? "engine-parity-out");
const RECORD = process.env["FOLIO_ENGINE_PARITY_RECORD"] === "1";

/**
 * Served by the playground from `tests/visual/fixtures`. Chosen to cover plain
 * prose, justified text, lists, tables, a right-to-left script, mixed scripts
 * and subset faces, with kerning enabled in `sample.docx` and left off in the
 * others.
 */
const FIXTURES = [
  "docx-editor-demo.docx", // prose, lists, tables
  "sample.docx", // short prose with a table, kerning on
  "markup-views-original.docx", // justified paragraph style
  "markup-views.docx", // justified paragraph style with markup
  "markup-views-final.docx", // justified paragraph style, markup resolved
  "tracked-insertion-boundary.docx", // one paragraph with an insertion boundary
  "cursive-face-change.docx", // Arabic (right-to-left) with a face change
  "font-subsets.docx", // diacritics from separate font subsets
  "mixed-script-font-measure.docx", // mixed scripts
] as const;

/** Load a document and wait for the first page to paint with fonts settled. */
async function openFixture(page: Page, fixture: string): Promise<void> {
  await page.goto(`/?file=${encodeURIComponent(fixture)}`);
  await page.waitForSelector(".layout-page .layout-line", { timeout: 60_000 });
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
}

/**
 * Read every page. Long documents virtualise their pages, so each page is
 * scrolled into view and given a moment to paint before its lines are read.
 */
async function recordDocument(page: Page, fixture: string): Promise<DocumentRecord> {
  const pageCount = await page.locator(".layout-page").count();
  const pages: PageRecord[] = [];
  const fonts = new Map<string, FontResolution>();
  for (let index = 0; index < pageCount; index++) {
    const pageEl = page.locator(".layout-page").nth(index);
    await pageEl.scrollIntoViewIfNeeded();
    const painted = await pageEl
      .locator(".layout-line")
      .first()
      .waitFor({ timeout: 5_000 })
      .then(
        () => true,
        () => false,
      );
    if (!painted) {
      pages.push({ rendered: false, lines: [] });
      continue;
    }
    const read = await pageEl.evaluate((el) => {
      const lines = [...el.querySelectorAll<HTMLElement>(".layout-line")].map((lineEl) => {
        const content = [...lineEl.querySelectorAll<HTMLElement>("[data-pm-start]")];
        const words = content
          .map((span) => span.textContent ?? "")
          .join("")
          .split(/\s+/u)
          .filter((word) => word.length > 0);
        const claimed = Number(lineEl.dataset["measuredWidth"]);
        return {
          first: words.at(0) ?? "",
          last: words.at(-1) ?? "",
          width: Number.isFinite(claimed) && lineEl.dataset["measuredWidth"] ? claimed : null,
        };
      });
      const stacks = [...el.querySelectorAll<HTMLElement>("[data-pm-start]")].map((span) => {
        const style = getComputedStyle(span);
        return { stack: style.fontFamily, weight: style.fontWeight };
      });
      return { lines, stacks };
    });
    pages.push({ rendered: true, lines: read.lines });
    for (const { stack, weight } of read.stacks) {
      const key = `${weight} ${stack}`;
      if (fonts.has(key)) continue;
      fonts.set(key, {
        stack,
        weight,
        loaded: await page.evaluate(([w, s]) => document.fonts.check(`${w} 16px ${s}`), [
          weight,
          stack,
        ] as const),
      });
    }
  }
  return { fixture, pages, fonts: [...fonts.values()] };
}

const readBaseline = (fixture: string): DocumentRecord | null => {
  const file = path.join(BASELINE_DIR, `${fixture}.json`);
  if (!fs.existsSync(file)) return null;
  // SAFETY: the file is written only by this spec's record mode.
  return JSON.parse(fs.readFileSync(file, "utf8")) as DocumentRecord;
};

const writeJson = (dir: string, name: string, value: unknown): void => {
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, name), `${JSON.stringify(value, null, 2)}\n`);
};

test.describe("layout under this engine vs the Chromium record", () => {
  for (const fixture of FIXTURES) {
    test(`records ${fixture}`, async ({ page, browserName }) => {
      test.setTimeout(240_000);
      await openFixture(page, fixture);
      const actual = await recordDocument(page, fixture);
      expect(actual.pages.length).toBeGreaterThan(0);

      if (RECORD) {
        expect(browserName).toBe("chromium");
        writeJson(BASELINE_DIR, `${fixture}.json`, actual);
        return;
      }

      const expected = readBaseline(fixture);
      const differences = expected === null ? [] : diffRecords(expected, actual);
      writeJson(OUT_DIR, `${fixture}.json`, {
        fixture,
        engine: browserName,
        status: expected === null ? "unrecorded" : "compared",
        pageCount: { expected: expected?.pages.length ?? null, actual: actual.pages.length },
        unrenderedPages: actual.pages.flatMap((p, i) => (p.rendered ? [] : [i + 1])),
        differences,
        fonts: { expected: expected?.fonts ?? null, actual: actual.fonts },
      });
    });
  }

  test("records whether fontKerning none changes measureText", async ({ page, browserName }) => {
    await openFixture(page, FIXTURES[0]);
    const result = await page.evaluate(async () => {
      const families = ["Arimo", "Carlito", "Tinos", "serif"];
      const text = "AVATAR To Yo We P. F, LT av r. Wa Wo Ya Tw Vo";
      const ctx = document.createElement("canvas").getContext("2d");
      if (ctx === null) throw new Error("no 2d context");
      const rows = [];
      for (const family of families) {
        await document.fonts.load(`32px ${family}`, text).catch(() => []);
        ctx.font = `32px ${family}`;
        ctx.fontKerning = "normal";
        const normal = ctx.measureText(text).width;
        ctx.fontKerning = "none";
        const none = ctx.measureText(text).width;
        rows.push({
          family,
          fontKerningSupported: "fontKerning" in ctx,
          loaded: document.fonts.check(`32px ${family}`),
          normal,
          none,
          delta: none - normal,
          changes: Math.abs(none - normal) > 0.01,
        });
      }
      return rows;
    });
    writeJson(OUT_DIR, "_kerning.json", { engine: browserName, text: "kerning pairs", result });
    expect(result.length).toBeGreaterThan(0);
  });
});
