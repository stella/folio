/**
 * Layout under a second browser engine: record where it differs, do not gate.
 *
 * Folio breaks lines and paginates itself, but glyph advances come from the
 * browser's canvas `measureText`. A different engine can therefore break a line
 * in a different place. This spec records, per fixture document, the page
 * count and for every painted line its first and last word and its measured
 * width. Under Chromium it writes the reference record into
 * `FOLIO_ENGINE_PARITY_REFERENCE` (default `engine-parity-ref`); under any other
 * engine it reads that record, taken on the same machine, and compares.
 *
 * It is report-only. A difference never fails a test: each document writes a
 * JSON file with its structured differences into `FOLIO_ENGINE_PARITY_OUT`
 * (default `engine-parity-out`). A test fails only on a harness problem: a
 * document that does not load, a page that never paints, a document with no
 * lines, or a missing reference file.
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

const REFERENCE_DIR = path.resolve(
  process.env["FOLIO_ENGINE_PARITY_REFERENCE"] ?? "engine-parity-ref",
);
const OUT_DIR = path.resolve(process.env["FOLIO_ENGINE_PARITY_OUT"] ?? "engine-parity-out");

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
    // Pages of long documents paint when they intersect the viewport; give
    // each a generous window and one more scroll before calling it unpainted.
    let painted = false;
    for (let attempt = 0; attempt < 2 && !painted; attempt++) {
      await pageEl.scrollIntoViewIfNeeded();
      painted = await pageEl
        .locator(".layout-line")
        .first()
        .waitFor({ timeout: 15_000 })
        .then(
          () => true,
          () => false,
        );
    }
    if (painted) {
      await page.evaluate(async () => {
        await document.fonts.ready;
      });
    }
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

const readReference = (fixture: string): DocumentRecord => {
  const file = path.join(REFERENCE_DIR, `${fixture}.json`);
  if (!fs.existsSync(file)) throw new Error(`reference record missing: ${file}`);
  // SAFETY: the file is written only by this spec under the reference engine.
  const record = JSON.parse(fs.readFileSync(file, "utf8")) as DocumentRecord & { status: string };
  if (record.status !== "recorded") throw new Error(`reference record unusable: ${file}`);
  return record;
};

const writeJson = (dir: string, name: string, value: unknown): void => {
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, name), `${JSON.stringify(value, null, 2)}\n`);
};

test.describe("layout under this engine vs the Chromium record", () => {
  for (const fixture of FIXTURES) {
    test(`records ${fixture}`, async ({ page, browserName }) => {
      test.setTimeout(240_000);
      const reference = browserName === "chromium";
      const dir = reference ? REFERENCE_DIR : OUT_DIR;
      // A problem with one document is written down and the run continues, so
      // the others are still recorded and compared. The summary step fails the
      // job when any document carries a harness error.
      try {
        await openFixture(page, fixture);
        const actual = await recordDocument(page, fixture);
        // Zero pages, a page that never painted, or no lines at all is a broken
        // harness, never "no differences".
        if (actual.pages.length === 0) throw new Error("document produced no pages");
        const unpainted = actual.pages.flatMap((p, i) => (p.rendered ? [] : [i + 1]));
        if (unpainted.length > 0) {
          throw new Error(`pages never painted: ${unpainted.join(", ")}`);
        }
        if (actual.pages.every((p) => p.lines.length === 0)) {
          throw new Error("document produced no lines");
        }

        if (reference) {
          writeJson(dir, `${fixture}.json`, { status: "recorded", ...actual });
          return;
        }

        const expected = readReference(fixture);
        writeJson(dir, `${fixture}.json`, {
          status: "compared",
          fixture,
          engine: browserName,
          pageCount: { expected: expected.pages.length, actual: actual.pages.length },
          differences: diffRecords(expected, actual),
          fonts: { expected: expected.fonts, actual: actual.fonts },
        });
      } catch (error) {
        writeJson(dir, `${fixture}.json`, {
          status: "harness-error",
          fixture,
          engine: browserName,
          message: error instanceof Error ? error.message : String(error),
        });
      }
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
    writeJson(browserName === "chromium" ? REFERENCE_DIR : OUT_DIR, "_kerning.json", {
      engine: browserName,
      text: "kerning pairs",
      result,
    });
    expect(result.length).toBeGreaterThan(0);
  });
});
