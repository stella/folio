/**
 * Measure/paint parity: does the width the engine decided match the width the
 * browser drew?
 *
 * Folio measures text with canvas and paints it as DOM text the browser shapes
 * itself. Two engines. `textMeasurementPolicy` keeps their inputs aligned by
 * convention, but convention is not a guarantee: nothing compared their
 * outputs, so the two could disagree and every unit test still pass. They did
 * disagree — a joiner that DOM layout honours and canvas ignores — and the
 * suite stayed green throughout.
 *
 * This spec is that missing comparison. Each painted line carries
 * `data-measured-width`, the measurer's claim; the painted extent comes from a
 * Range over the line's own content. Both sides are read from the same browser
 * with the same fonts, so unlike a screenshot baseline this is environment
 * independent and safe to gate CI on.
 *
 * A divergence here is a real defect, not a rendering nuance: line breaking and
 * pagination are decided on the measured number, so wherever it disagrees with
 * what is drawn, the page is laid out against a fiction.
 */

import { test, expect, type Page } from "@playwright/test";

/** Per-line comparison, in CSS pixels at zoom 1. */
type LineParity = {
  index: number;
  measured: number;
  painted: number;
  delta: number;
  text: string;
  /** Whether the line carries joiner furniture, i.e. a repaired face change. */
  hasJoiner: boolean;
};

/**
 * Sub-pixel slack. Canvas advances and layout advances round independently, and
 * a line accumulates one rounding per run, so the budget scales with run count
 * rather than being a flat number that silently absorbs a real regression.
 *
 * Counted over CONTENT runs only. Counting every child would let each joiner
 * span, list marker and break marker buy another half-pixel of slack, so a
 * repaired line would be judged more loosely than the plain line beside it.
 */
const PER_RUN_TOLERANCE_PX = 0.5;

/**
 * How far a repaired cursive line may diverge, as a fraction of its width.
 *
 * Wider than the per-run rounding budget because this is not rounding: the
 * measurer cannot see the joiners at all, so it measures unjoined forms while
 * the browser paints joined ones. Sized to bound the observed disagreement,
 * not to hide it, and it shrinks to nothing once measurement is shaped.
 */
const REPAIRED_LINE_BUDGET_RATIO = 0.02;

/**
 * Read every painted line and compare the measurer's claim with the drawn
 * extent. Runs entirely in the page: both numbers must come from one layout
 * pass, or a resize between them would be read as a divergence.
 */
async function collectLineParity(page: Page): Promise<LineParity[]> {
  return page.evaluate(() => {
    const lines = [...document.querySelectorAll<HTMLElement>(".layout-line")];
    const results: {
      index: number;
      measured: number;
      painted: number;
      delta: number;
      text: string;
      hasJoiner: boolean;
      runCount: number;
    }[] = [];

    lines.forEach((lineEl, index) => {
      const claimed = lineEl.dataset["measuredWidth"];
      if (claimed === undefined) return;
      const measured = Number(claimed);
      if (!Number.isFinite(measured) || measured <= 0) return;

      // Compare like with like: `MeasuredLine.width` covers the line's CONTENT
      // runs. Page furniture painted into the same element — the list marker on
      // a first line, the zero-width break marker, joiner spans — is positioned
      // by other means and is not in that number, so a Range over the whole
      // element would read the marker's width as a measurement error. Content
      // is exactly what carries pm positions.
      const content = [...lineEl.querySelectorAll<HTMLElement>("[data-pm-start]")];
      if (content.length === 0) return;
      const text = content.map((el) => el.textContent ?? "").join("");
      if (text.trim().length === 0) return;

      const range = document.createRange();
      // SAFETY: length checked above.
      range.setStartBefore(content[0]!);

      // Trailing whitespace at a soft wrap hangs past the line, exactly as Word
      // renders it, and the measurer excludes it from the line width on purpose
      // (`trimTrailingSpacesAndTabs`). End the range at the last non-whitespace
      // character so both sides describe the same span of text.
      let ended = false;
      for (let i = content.length - 1; i >= 0 && !ended; i--) {
        // SAFETY: i is inside the array bounds.
        const span = content[i]!;
        const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
        const textNodes: Text[] = [];
        let node = walker.nextNode();
        while (node !== null) {
          textNodes.push(node as Text);
          node = walker.nextNode();
        }
        for (let j = textNodes.length - 1; j >= 0; j--) {
          // SAFETY: j is inside the array bounds.
          const textNode = textNodes[j]!;
          const trimmed = (textNode.data ?? "").replace(/[\s\u200B]+$/u, "");
          if (trimmed.length === 0) continue;
          range.setEnd(textNode, trimmed.length);
          ended = true;
          break;
        }
      }
      if (!ended) return;

      const painted = range.getBoundingClientRect().width;
      range.detach();
      if (painted <= 0) return;

      results.push({
        index,
        measured,
        painted,
        delta: painted - measured,
        text: text.slice(0, 60),
        hasJoiner: lineEl.querySelector("[data-docx-joiner]") !== null,
        runCount: content.length,
      });
    });
    return results;
  });
}

/** Load a document into the playground and wait for the first page to paint. */
async function openFixture(page: Page, fixture: string): Promise<void> {
  await page.goto(`/?file=${encodeURIComponent(fixture)}`);
  await page.waitForSelector(".layout-page .layout-line", { timeout: 30_000 });
  // Web fonts change advances, so a comparison taken mid-swap is meaningless.
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
}

const CURSIVE_FIXTURE = "cursive-face-change.docx";
const GENERAL_FIXTURE = "docx-editor-demo.docx";

test.describe("measure/paint parity", () => {
  test("the measurer's width matches the painted width on every line", async ({ page }) => {
    await openFixture(page, GENERAL_FIXTURE);
    const lines = await collectLineParity(page);

    // A floor, not a `> 0`: if the content selector ever stops matching, this
    // spec must fail rather than quietly compare a handful of lines.
    expect(lines.length).toBeGreaterThanOrEqual(100);

    const offenders = lines.filter(
      (line) => Math.abs(line.delta) > PER_RUN_TOLERANCE_PX * Math.max(1, line.runCount),
    );
    expect(
      offenders.map(
        (o) =>
          `line ${o.index} "${o.text}": measured ${o.measured.toFixed(2)}px, painted ${o.painted.toFixed(2)}px (${o.delta > 0 ? "+" : ""}${o.delta.toFixed(2)}px)`,
      ),
    ).toEqual([]);
  });

  test("a repaired cursive line stays within its measurement budget", async ({ page }) => {
    // The known residual, pinned as a number rather than left as a comment.
    // Canvas cannot see the joiners the painter inserts to keep a cursive word
    // connected across a face change, so the measurer reserves the unjoined
    // width and the painted width differs.
    //
    // The direction is a property of the FONT, not of the repair: joined forms
    // are narrower than isolated ones in some faces and wider in others. An
    // earlier version of this assertion encoded "the measurer always
    // over-reserves", which held for the macOS fallback face and failed on
    // Linux. The fixture now pins the complex-script font to one folio ships as
    // a webfont, so the relationship is the same everywhere and can be asserted
    // at all.
    await openFixture(page, CURSIVE_FIXTURE);
    const lines = await collectLineParity(page);
    const repaired = lines.filter((line) => line.hasJoiner);

    // The fixture exists to produce these; zero of them means it stopped
    // exercising the path and the assertion below is vacuous.
    expect(repaired.length).toBeGreaterThan(0);

    // Exact, because the fixture is exact. Paragraph 1 splits four words with a
    // bold second half, but only THREE of those splits severed a connection:
    // the fourth splits after alef, which joins backwards only, so no joiner
    // belongs there. Three boundaries x two joiners = six. Paragraph 2 splits
    // the same words with colour only (same face, browser already joins) and
    // paragraph 3 is Latin, so both must contribute nothing. A higher count
    // means a false positive; a lower one means a repair stopped happening.
    const joinerSpans = await page.evaluate(
      () => document.querySelectorAll("[data-docx-joiner]").length,
    );
    expect(joinerSpans).toBe(6);

    // Bound the MAGNITUDE, not the direction. Whether joined forms are wider or
    // narrower than isolated ones is a property of the face: unpinned, this
    // divergence was -2.7px on the macOS Arabic fallback and +2.5px on the Linux
    // one. So the residual is a bounded disagreement that can push a line either
    // way, not the safe over-reservation an earlier version of this claimed.
    const budget = (line: LineParity): number =>
      Math.max(3, line.measured * REPAIRED_LINE_BUDGET_RATIO);
    expect(
      repaired
        .filter((line) => Math.abs(line.delta) > budget(line))
        .map(
          (line) =>
            `measured ${line.measured.toFixed(2)}px, painted ${line.painted.toFixed(2)}px, over budget ${budget(line).toFixed(2)}px`,
        ),
    ).toEqual([]);
  });

  test("joiner furniture is invisible to the document's text", async ({ page }) => {
    // The joiners must not reach anything that reads the document: they carry no
    // pm positions, so span-to-position mapping and text extraction skip them.
    await openFixture(page, CURSIVE_FIXTURE);
    const leaked = await page.evaluate(
      () =>
        [...document.querySelectorAll<HTMLElement>("[data-docx-joiner]")].filter(
          (el) => el.dataset["pmStart"] !== undefined || el.dataset["pmEnd"] !== undefined,
        ).length,
    );
    expect(leaked).toBe(0);
  });
});
