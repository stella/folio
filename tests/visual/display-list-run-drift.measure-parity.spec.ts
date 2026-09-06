/**
 * Does the display-list renderer paint runs where, and as wide as, the display
 * list says?
 *
 * This is the arrangement the editor actually uses, and it is the only one
 * worth gating: the display list is built by the canvas measure provider
 * inside the same browser that then paints it, so both numbers come from one
 * layout pass and different system fonts move neither relative to the other.
 * That is what makes it environment-independent in the way a screenshot
 * baseline is not, and the same reason `measure-parity.spec.ts` is safe to
 * gate on.
 *
 * Both the run's origin and its extent are gated. The extent used to be
 * reported only, because `advancesPx` folded in letter spacing, a horizontal
 * scale, justification and small capitals, and a shaped text node cannot
 * reproduce those from its text alone. The list now carries each of them
 * beside the advances, so the backend reapplies them and the extent is a
 * property the renderer either honours or does not.
 *
 * The paint-equivalence harness measures a different arrangement again, a list
 * built from font-table advances and painted by a shaper that kerns, and
 * reports that as a diagnostic. A number from there says nothing about the
 * editor.
 */

import { expect, test, type Page } from "@playwright/test";

/**
 * How far a run's painted geometry may sit from its declared geometry.
 *
 * A browser snaps each layout value to 1/64 px (0.0156), and both an origin and
 * an extent are set directly rather than accumulated, so a correctly painted
 * run can miss by one quantum and no more. Two quanta bounds the observed
 * disagreement rather than hiding it; a run painted from inline flow, or one
 * whose spacing the backend did not reapply, misses by whole pixels.
 */
const TOLERANCE_PX = 0.032;

type RunDrift = {
  readonly text: string;
  readonly declaredPx: number;
  readonly paintedPx: number;
  /** Painted extent against what the list declares the run occupies. */
  readonly extentDeltaPx: number;
  /** Painted left edge against declared. */
  readonly originDeltaPx: number;
};

/**
 * Read every painted run's declared geometry against the geometry the browser
 * gave it. Both come from one layout pass; a resize between them would read as
 * drift.
 */
const collectRunDrift = (page: Page): Promise<RunDrift[]> =>
  page.evaluate(() => {
    /** A rotated page cannot be compared; a run's own horizontal scale can. */
    const comparable = (element: HTMLElement): boolean => {
      for (let node = element.parentElement; node; node = node.parentElement) {
        const { transform } = getComputedStyle(node);
        if (transform !== "none" && transform !== "") {
          return false;
        }
      }
      const own = new DOMMatrixReadOnly(getComputedStyle(element).transform);
      return own.b === 0 && own.c === 0;
    };

    const drifts: RunDrift[] = [];
    for (const run of document.querySelectorAll<HTMLElement>("[data-advance-sum]")) {
      if (!comparable(run)) {
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
      const style = getComputedStyle(run);
      // CSS puts a letter-space after the run's last character; OOXML counts
      // the gaps between characters, so the declared width has one fewer. The
      // difference is the whole of it, and it scales with the run.
      const letterSpacingPx = Number.parseFloat(style.letterSpacing);
      const scale = new DOMMatrixReadOnly(style.transform).a;
      const trailingPx = Number.isFinite(letterSpacingPx) ? letterSpacingPx * scale : 0;
      const declaredLeft = Number.parseFloat(run.style.left);
      const parent = run.offsetParent;
      const paintedLeft =
        run.getBoundingClientRect().left -
        (parent === null ? 0 : parent.getBoundingClientRect().left);
      drifts.push({
        text: run.textContent ?? "",
        declaredPx,
        paintedPx,
        extentDeltaPx: Math.abs(paintedPx - (declaredPx + trailingPx)),
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

const report = (label: string, drifts: readonly RunDrift[]): void => {
  const extents = drifts.map((drift) => drift.extentDeltaPx);
  const origins = drifts.map((drift) => drift.originDeltaPx);
  // eslint-disable-next-line no-console -- the measurement is the deliverable
  console.log(
    `${label}: ${String(drifts.length)} runs | extent px median ` +
      `${percentile(extents, 0.5).toFixed(3)} p90 ${percentile(extents, 0.9).toFixed(3)} ` +
      `worst ${percentile(extents, 1).toFixed(3)} | origin px p90 ` +
      `${percentile(origins, 0.9).toFixed(3)} worst ${percentile(origins, 1).toFixed(3)}`,
  );
};

const describe = (drift: RunDrift, kind: "origin" | "extent", deltaPx: number) =>
  `"${drift.text.slice(0, 40)}": ${kind} off by ${deltaPx.toFixed(4)}px ` +
  `(declared ${drift.declaredPx.toFixed(3)}, painted ${drift.paintedPx.toFixed(3)})`;

const expectGeometryHonoured = (drifts: readonly RunDrift[]): void => {
  expect(drifts.length, "no glyph runs were painted").toBeGreaterThan(0);
  expect(
    drifts
      .filter((drift) => drift.originDeltaPx > TOLERANCE_PX)
      .map((drift) => describe(drift, "origin", drift.originDeltaPx)),
  ).toEqual([]);
  expect(
    drifts
      .filter((drift) => drift.extentDeltaPx > TOLERANCE_PX)
      .map((drift) => describe(drift, "extent", drift.extentDeltaPx)),
  ).toEqual([]);
};

test.describe("display-list run drift", () => {
  test("a run is painted where and as wide as the display list declares", async ({ page }) => {
    await openWithDisplayListRenderer(page, "sample.docx");
    const drifts = await collectRunDrift(page);

    report("sample.docx", drifts);
    expectGeometryHonoured(drifts);
  });

  test("a multi-page document honours every run's geometry", async ({ page }) => {
    // Past the virtualization threshold, so the incremental path is measured
    // rather than the eager small-document path.
    await openWithDisplayListRenderer(page, "podily-bps.docx");
    const drifts = await collectRunDrift(page);

    report("podily-bps.docx", drifts);
    expectGeometryHonoured(drifts);
  });
});
