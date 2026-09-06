/**
 * What each renderer costs to put a page on the screen.
 *
 * Both renderers are handed the same layout and paint into the same page
 * container, so the difference between them is the painting: the one that walks
 * `Layout` directly against the one that builds the paint IR from it and paints
 * from that. The IR renderer does strictly more work per run (it builds the
 * list, then walks it), and the question this answers is how much more.
 *
 * Two numbers per repaint, both read from the same browser in the same session
 * so neither depends on the machine:
 *
 * - **Ours**: the `render-pages` phase the layout pipeline already reports,
 *   which for the IR renderer includes building the list for the whole layout.
 * - **The frame**: from the moment the repaint is asked for to the frame that
 *   shows it, which is ours plus the browser's own style, layout and paint.
 *
 * A document past the virtualization threshold paints a window of pages rather
 * than all of them, so both numbers are reported per painted page: what a
 * reader waits for is a screen, not a document.
 */

import { expect, test, type Page } from "@playwright/test";

// Repeated repaints of a 91-page document for both renderers.
test.setTimeout(180_000);

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import type { LayoutInstrumentation } from "../../packages/core/src/layout-engine/layoutInstrumentation";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
  var __folioRenderPagesMs: number[] | undefined;
}

/** 91 pages: past the virtualization threshold, and long enough to be work. */
const FIXTURE = "performance-1500-paragraphs.docx";

/** Warm both rendering paths before collecting the steady-state samples. */
const WARMUP_REPAINTS = 3;

/** Sub-millisecond page costs need enough samples to resist scheduler noise. */
const MEASURED_REPAINTS = 11;

/**
 * How much more per painted page the IR renderer may cost.
 *
 * It builds a page's primitives and then walks them, where the other renderer
 * walks the layout once, so it need not be cheaper; measured, the two are
 * within a few percent of each other. The budget bounds a regression rather
 * than the current difference: building the whole document's list to paint a
 * screenful of it, which is what this renderer used to do, costs about ninety
 * times a page's own work on a ninety-page document, and any return to work
 * proportional to the document rather than the screen misses this by far more
 * than the noise between two runs.
 */
const COST_RATIO_BUDGET = 1.5;

type RepaintCost = {
  /** The pipeline's own `render-pages` phase, in ms. */
  readonly ourMs: number;
  /** Request to painted frame, in ms: ours plus the browser's own work. */
  readonly frameMs: number;
  readonly paintedPages: number;
};

const installInstrumentation = async (page: Page): Promise<void> => {
  await page.addInitScript(() => {
    globalThis.__folioRenderPagesMs = [];
    globalThis.__folioLayoutInstrumentation = {
      onLayoutPhase: ({ durationMs, phase }) => {
        if (phase === "render-pages") {
          globalThis.__folioRenderPagesMs?.push(durationMs);
        }
      },
    };
  });
};

const open = async (page: Page, renderer: string): Promise<void> => {
  // A ninety-page document takes a while to lay out for the first time, and
  // this spec opens it twice in one test.
  await page.goto(`/?file=${FIXTURE}&pageRenderer=${renderer}`, { timeout: 60_000 });
  await page.waitForSelector('[data-testid="folio-editor"]');
  await page.waitForFunction(() => document.querySelectorAll(".layout-page").length > 4);
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
};

/**
 * One full repaint, with both numbers read from the same run.
 *
 * Driven by loading the document the editor already holds: that lays it out
 * again from scratch and repaints every page on screen, which is the work being
 * compared. A zoom change would not do (it scales pages already painted, so it
 * measures a CSS transform), and typing would not either (it needs a caret,
 * which is placed by machinery this measurement is not about).
 */
const repaint = (page: Page): Promise<RepaintCost> =>
  page.evaluate(async () => {
    const editor = globalThis.__folioPlayground?.getEditorRef();
    const document_ = editor?.getDocument();
    if (!editor || !document_) {
      throw new Error("the playground exposed no loaded document");
    }
    const before = globalThis.__folioRenderPagesMs?.length ?? 0;
    const startedAt = performance.now();
    editor.loadDocument(document_);
    // The pipeline runs off that call rather than in it, so wait until it has
    // recorded the phase, then for the frame that shows the result.
    const deadline = performance.now() + 20_000;
    while (
      (globalThis.__folioRenderPagesMs?.length ?? 0) === before &&
      performance.now() < deadline
    ) {
      await new Promise<void>((resolve) => {
        setTimeout(resolve, 4);
      });
    }
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          resolve();
        });
      });
    });
    const frameMs = performance.now() - startedAt;
    const recorded = globalThis.__folioRenderPagesMs ?? [];
    if (recorded.length === before) {
      throw new Error("the repaint produced no render-pages timing before the deadline");
    }
    const ourMs = recorded.slice(before).reduce((total, value) => total + value, 0);
    // A page the container has not filled yet paints nothing, and counting it
    // would flatter whichever renderer happened to be given fewer.
    const paintedPages = [...document.querySelectorAll(".layout-page")].filter(
      (element) => element.childElementCount > 0,
    ).length;
    return { ourMs, frameMs, paintedPages };
  });

const median = (values: readonly number[]): number => {
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
};

const perPage = (costs: readonly RepaintCost[], read: (cost: RepaintCost) => number): number =>
  median(costs.map((cost) => read(cost) / Math.max(1, cost.paintedPages)));

const measure = async (page: Page, renderer: string): Promise<readonly RepaintCost[]> => {
  await open(page, renderer);
  const costs: RepaintCost[] = [];
  for (let index = 0; index < WARMUP_REPAINTS + MEASURED_REPAINTS; index += 1) {
    costs.push(await repaint(page));
  }
  return costs.slice(WARMUP_REPAINTS);
};

test.describe("what each renderer costs per page", () => {
  test("the display list paints a page within a bounded factor of the painter", async ({
    page,
  }) => {
    await installInstrumentation(page);

    const legacy = await measure(page, "legacy");
    const displayList = await measure(page, "display-list");

    const report = (label: string, costs: readonly RepaintCost[]) =>
      `${label}: ${costs.length} repaints over ${String(costs.at(0)?.paintedPages ?? 0)} painted pages | ` +
      `render-pages ${perPage(costs, (cost) => cost.ourMs).toFixed(3)} ms/page | ` +
      `to painted frame ${perPage(costs, (cost) => cost.frameMs).toFixed(3)} ms/page`;
    // eslint-disable-next-line no-console -- the measurement is the deliverable
    console.log(`${report("legacy", legacy)}\n${report("display-list", displayList)}`);

    expect(legacy.length, "no repaints measured for the painter").toBeGreaterThan(0);
    expect(displayList.length, "no repaints measured for the display list").toBeGreaterThan(0);
    expect(
      legacy.every((cost) => cost.paintedPages > 0),
      "the painter painted no pages",
    ).toBe(true);
    expect(
      displayList.every((cost) => cost.paintedPages > 0),
      "the display list painted no pages",
    ).toBe(true);

    const ratio =
      perPage(displayList, (cost) => cost.ourMs) /
      Math.max(
        0.001,
        perPage(legacy, (cost) => cost.ourMs),
      );
    expect({ ratio: Number(ratio.toFixed(2)), budget: COST_RATIO_BUDGET }).toEqual({
      ratio: Number(Math.min(ratio, COST_RATIO_BUDGET).toFixed(2)),
      budget: COST_RATIO_BUDGET,
    });
  });
});
