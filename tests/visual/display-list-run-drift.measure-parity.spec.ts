/**
 * Does the display-list renderer paint runs where the display list says?
 *
 * This is the arrangement the editor actually uses, and it is the only one
 * worth gating: the display list is built by the canvas measure provider
 * inside the same browser that then paints it, so both numbers come from one
 * layout pass and different system fonts move neither relative to the other.
 * That is what makes it environment-independent in the way a screenshot
 * baseline is not, and the same reason `measure-parity.spec.ts` is safe to
 * gate on.
 *
 * What is gated is the run's *origin*, because that is what the display list
 * decides. Where each glyph lands inside a run is the shaper's business, and
 * so is the run's painted extent: `advancesPx` folds in letter spacing,
 * horizontal scale and justification adjustments, and one shaped text node
 * cannot reproduce those from its text alone. The extent difference is
 * therefore measured and reported here rather than asserted. Closing it would
 * mean carrying those adjustments in the paint IR beside the advances instead
 * of folded into them, so a backend could reapply them.
 *
 * The paint-equivalence harness measures a different arrangement again, a list
 * built from font-table advances and painted by a shaper that kerns, and
 * reports that as a diagnostic. A number from there says nothing about the
 * editor.
 */

import { expect, test, type Page } from "@playwright/test";

/**
 * How far a run's painted left edge may sit from its declared one.
 *
 * A browser snaps each layout value to 1/64 px (0.0156), and an origin is set
 * directly rather than accumulated, so a correctly placed run can miss by one
 * quantum and no more. Two quanta bounds the observed disagreement rather than
 * hiding it; a run placed from inline flow instead of from the display list
 * would miss by whole pixels and fail here.
 */
const ORIGIN_TOLERANCE_PX = 0.032;

type RunDrift = {
  readonly text: string;
  readonly declaredPx: number;
  readonly paintedPx: number;
  /** Painted extent against declared: reported, not gated. */
  readonly extentDeltaPx: number;
  /** Painted left edge against declared: gated. */
  readonly originDeltaPx: number;
};

/**
 * Read every painted run's declared geometry against the geometry the browser
 * gave it. Both come from one layout pass; a resize between them would read as
 * drift.
 */
const collectRunDrift = (page: Page): Promise<RunDrift[]> =>
  page.evaluate(() => {
    const axisAligned = (element: HTMLElement): boolean => {
      for (let node: HTMLElement | null = element; node; node = node.parentElement) {
        const { transform } = getComputedStyle(node);
        if (transform !== "none" && transform !== "") {
          return false;
        }
      }
      return true;
    };

    const drifts: RunDrift[] = [];
    for (const run of document.querySelectorAll<HTMLElement>("[data-advance-sum]")) {
      if (!axisAligned(run)) {
        continue;
      }
      const declaredPx = Number(run.dataset["advanceSum"]);
      if (!Number.isFinite(declaredPx) || declaredPx <= 0) {
        continue;
      }
      const range = document.createRange();
      range.selectNodeContents(run);
      const paintedPx = range.getBoundingClientRect().width;
      range.detach();
      if (paintedPx <= 0) {
        continue;
      }
      const declaredLeft = Number.parseFloat(run.style.left);
      const parent = run.offsetParent;
      const paintedLeft =
        run.getBoundingClientRect().left -
        (parent === null ? 0 : parent.getBoundingClientRect().left);
      drifts.push({
        text: run.textContent ?? "",
        declaredPx,
        paintedPx,
        extentDeltaPx: Math.abs(paintedPx - declaredPx),
        originDeltaPx: Number.isFinite(declaredLeft) ? Math.abs(paintedLeft - declaredLeft) : 0,
      });
    }
    return drifts;
  });

const openWithDisplayListRenderer = async (page: Page, fixture: string): Promise<void> => {
  await page.goto(`/?file=${fixture}&pageRenderer=display-list`);
  await page.waitForSelector("[data-advance-sum]", { timeout: 30_000 });
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
};

const percentile = (values: readonly number[], fraction: number): number => {
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * fraction))] ?? 0;
};

const reportExtents = (label: string, drifts: readonly RunDrift[]): void => {
  const deltas = drifts.map((drift) => drift.extentDeltaPx);
  const ratios = drifts.map((drift) => drift.extentDeltaPx / drift.declaredPx);
  // eslint-disable-next-line no-console -- the measurement is the deliverable
  console.log(
    `${label}: ${String(drifts.length)} runs | extent px median ` +
      `${percentile(deltas, 0.5).toFixed(3)} p90 ${percentile(deltas, 0.9).toFixed(3)} ` +
      `worst ${percentile(deltas, 1).toFixed(3)} | extent ratio median ` +
      `${(percentile(ratios, 0.5) * 100).toFixed(3)}% p90 ${(percentile(ratios, 0.9) * 100).toFixed(3)}%`,
  );
};

const describeOrigin = (drift: RunDrift) =>
  `"${drift.text.slice(0, 40)}": origin off by ${drift.originDeltaPx.toFixed(4)}px`;

const expectOriginsHonoured = (drifts: readonly RunDrift[]): void => {
  expect(drifts.length, "no glyph runs were painted").toBeGreaterThan(0);
  expect(
    drifts.filter((drift) => drift.originDeltaPx > ORIGIN_TOLERANCE_PX).map(describeOrigin),
  ).toEqual([]);
};

test.describe("display-list run drift", () => {
  test("a run is painted at the origin the display list declares", async ({ page }) => {
    await openWithDisplayListRenderer(page, "sample.docx");
    const drifts = await collectRunDrift(page);

    reportExtents("sample.docx", drifts);
    expectOriginsHonoured(drifts);
  });

  test("a multi-page document honours every run's origin", async ({ page }) => {
    // Past the virtualization threshold, so the incremental path is measured
    // rather than the eager small-document path.
    await openWithDisplayListRenderer(page, "podily-bps.docx");
    const drifts = await collectRunDrift(page);

    reportExtents("podily-bps.docx", drifts);
    expectOriginsHonoured(drifts);
  });
});
