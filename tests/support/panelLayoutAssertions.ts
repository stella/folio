/**
 * Layout assertions for the editor's side panels (outline, comments), shared
 * by the playground spec and the VS Code webview spec so both hosts are held
 * to the same oracle: at every width, the page, the outline and the comments
 * occupy disjoint columns, and a panel that does not fit is a drawer that
 * opens over the page and closes on Escape or an outside press.
 */

import { expect, type Locator, type Page } from "@playwright/test";

import {
  PANEL_METRICS,
  type CommentsPresentation,
  type OutlinePresentation,
  type PanelLayoutTier,
} from "@stll/folio-core/panel-layout";
import type { PanelLayoutReview } from "./panelLayoutDocument";

export type PanelState = {
  tier: PanelLayoutTier;
  outline: OutlinePresentation;
  comments: CommentsPresentation;
};

/** Closing comments as a drawer also closes their requested panel visibility. */
export const afterPanelDrawersDismissed = (state: PanelState): PanelState => {
  if (state.comments !== "drawer") {
    return state;
  }
  let tier: PanelLayoutTier = "wide";
  if (state.outline === "drawer") {
    tier = "narrow";
  } else if (state.outline === "rail") {
    tier = "medium";
  }
  return { ...state, tier, comments: "hidden" };
};

export type PanelLayoutCase = {
  width: number;
  review: PanelLayoutReview;
  expected: PanelState;
};

/**
 * The tiers a US Letter page at 100% lands in. The widths sit well clear of
 * the derived thresholds (page + panels), so a scrollbar's width either way
 * cannot move a case.
 */
export const PANEL_LAYOUT_CASES: readonly PanelLayoutCase[] = [
  {
    width: 480,
    review: "none",
    expected: { tier: "narrow", outline: "drawer", comments: "hidden" },
  },
  {
    width: 480,
    review: "comment-and-changes",
    expected: { tier: "narrow", outline: "drawer", comments: "drawer" },
  },
  {
    width: 800,
    review: "none",
    expected: { tier: "narrow", outline: "drawer", comments: "hidden" },
  },
  {
    width: 800,
    review: "comment-and-changes",
    expected: { tier: "narrow", outline: "drawer", comments: "drawer" },
  },
  {
    width: 1100,
    review: "none",
    expected: { tier: "medium", outline: "rail", comments: "hidden" },
  },
  {
    width: 1100,
    review: "comment-and-changes",
    expected: { tier: "narrow", outline: "rail", comments: "drawer" },
  },
  {
    width: 1500,
    review: "none",
    expected: { tier: "wide", outline: "column", comments: "hidden" },
  },
  {
    width: 1500,
    review: "comment-and-changes",
    expected: { tier: "wide", outline: "column", comments: "column" },
  },
];

const PANELS_ROW = "[data-folio-panel-tier]";

/** The presentation the editor reports on its panels row, as written there. */
export type ReportedPanelState = Record<keyof PanelState, string | undefined>;

export const readPanelState = (page: Page): Promise<ReportedPanelState | null> =>
  page.evaluate((selector) => {
    const row = document.querySelector<HTMLElement>(selector);
    if (!row) return null;
    return {
      tier: row.dataset["folioPanelTier"],
      outline: row.dataset["folioOutline"],
      comments: row.dataset["folioComments"],
    };
  }, PANELS_ROW);

/**
 * Wait for a painted page as well as panel state. Heading collection and
 * the host's loaded signal can precede the initial font-ready layout.
 * Reviewed documents also wait for their comments to auto-open.
 */
export const waitForPanels = async (page: Page, review: PanelLayoutReview): Promise<void> => {
  await expect(page.locator(".layout-page").first()).toBeVisible();
  await expect
    .poll(async () => {
      const state = await readPanelState(page);
      if (!state || state.outline === "none") return false;
      return review !== "comment-and-changes" || state.comments !== "hidden";
    })
    .toBe(true);
};

type Interval = { left: number; right: number };
type PanelBox = Interval & { layoutWidth: number };

type PanelBoxes = {
  pages: Interval[];
  outline: PanelBox | null;
  comments: PanelBox | null;
  cards: Interval[];
  viewport: Interval;
  horizontalOverflow: number;
};

const readPanelBoxes = (page: Page): Promise<PanelBoxes> =>
  page.evaluate(() => {
    const interval = (element: Element | null) => {
      if (!element) return null;
      const { left, right, width } = element.getBoundingClientRect();
      return width > 0 ? { left, right } : null;
    };
    const panelBox = (element: HTMLElement | null): PanelBox | null => {
      const bounds = interval(element);
      return bounds ? { ...bounds, layoutWidth: element.offsetWidth } : null;
    };
    const present = <T>(value: T | null): value is T => value !== null;
    const scroll = document.querySelector<HTMLElement>("[data-folio-scroll]");
    const scrollBox = scroll?.getBoundingClientRect();
    return {
      pages: [...document.querySelectorAll(".layout-page")].map(interval).filter(present),
      outline: panelBox(
        document.querySelector<HTMLElement>(
          '[data-testid="folio-outline"]:not([data-folio-outline-surface="drawer"])',
        ),
      ),
      comments: panelBox(
        document.querySelector<HTMLElement>('[data-folio-comments-surface="column"]'),
      ),
      cards: [
        ...document.querySelectorAll('[data-folio-comments-surface="column"] .docx-comment-card'),
      ]
        .map(interval)
        .filter(present),
      viewport: {
        left: scrollBox?.left ?? 0,
        right: (scrollBox?.left ?? 0) + (scroll?.clientWidth ?? 0),
      },
      horizontalOverflow: scroll ? scroll.scrollWidth - scroll.clientWidth : 0,
    };
  });

/** Sub-pixel rounding between two boxes laid side by side. */
const EPSILON = 0.5;

const overlaps = (a: Interval, b: Interval) =>
  a.left < b.right - EPSILON && b.left < a.right - EPSILON;

/**
 * The oracle: every in-flow panel (outline column or rail, comments column
 * and its cards) shares no horizontal span with a page or with each other,
 * and when no panel is a drawer the page fits without horizontal scrolling.
 */
export const expectPanelsDoNotOverlap = async (page: Page, state: PanelState): Promise<void> => {
  const boxes = await readPanelBoxes(page);
  expect(boxes.pages.length).toBeGreaterThan(0);
  const ruler = page.getByTestId("folio-horizontal-ruler");
  if ((await ruler.count()) > 0) await expectOpaque(ruler, "the ruler");
  const panels: [string, Interval][] = [];
  if (boxes.outline) panels.push(["outline", boxes.outline]);
  if (boxes.comments) panels.push(["comments", boxes.comments]);
  for (const card of boxes.cards) panels.push(["comment card", card]);

  expect(boxes.outline !== null).toBe(state.outline === "column" || state.outline === "rail");
  expect(boxes.comments !== null).toBe(state.comments === "column");
  if (boxes.outline) {
    const width =
      state.outline === "rail" ? PANEL_METRICS.outlineRailWidth : PANEL_METRICS.outlineColumnWidth;
    expect(boxes.outline.layoutWidth).toBeCloseTo(width, 0);
  }
  if (boxes.comments) {
    expect(boxes.comments.layoutWidth).toBeCloseTo(PANEL_METRICS.commentsWidth, 0);
  }

  for (const [name, panel] of panels) {
    for (const pageBox of boxes.pages) {
      expect(overlaps(panel, pageBox), `${name} overlaps a page`).toBe(false);
    }
  }
  if (boxes.outline && boxes.comments) {
    expect(overlaps(boxes.outline, boxes.comments), "outline overlaps comments").toBe(false);
  }
  if (state.tier !== "narrow") {
    expect(boxes.horizontalOverflow, "the page scrolls sideways").toBeLessThanOrEqual(1);
    for (const pageBox of boxes.pages) {
      expect(pageBox.left).toBeGreaterThanOrEqual(boxes.viewport.left - EPSILON);
      expect(pageBox.right).toBeLessThanOrEqual(boxes.viewport.right + EPSILON);
    }
  }
};

/**
 * Chrome that pages scroll under (the sticky ruler, a comments drawer) must be
 * opaque, or page text shows through it: its background colour, under any
 * image layers, has full alpha.
 */
export const expectOpaque = async (locator: Locator, name: string): Promise<void> => {
  const alpha = await locator.evaluate((element) => {
    const probe = document.createElement("canvas").getContext("2d");
    if (!probe) return 0;
    probe.fillStyle = getComputedStyle(element).backgroundColor;
    probe.fillRect(0, 0, 1, 1);
    return probe.getImageData(0, 0, 1, 1).data[3] ?? 0;
  });
  expect(alpha, `${name} lets the page show through`).toBe(255);
};

const focusIsWithin = (locator: Locator) =>
  locator.evaluate((element) => element.contains(document.activeElement));

/**
 * A drawer opened by `opener`: it opens over the page with a scrim and takes
 * focus; opened from the keyboard, Escape closes it and hands focus back to
 * the opener; opened with the pointer, a press outside closes it.
 */
export const expectDrawerCycle = async (
  page: Page,
  { opener, drawer }: { opener: Locator; drawer: Locator },
): Promise<void> => {
  const scrim = page.getByTestId("folio-panel-scrim");

  await opener.focus();
  await page.keyboard.press("Enter");
  await expect(drawer).toBeVisible();
  await expect(scrim).toBeVisible();
  await expect.poll(() => focusIsWithin(drawer)).toBe(true);
  await page.keyboard.press("Escape");
  await expect(drawer).toBeHidden();
  await expect(scrim).toBeHidden();
  await expect.poll(() => focusIsWithin(opener)).toBe(true);

  await opener.click();
  await expect(drawer).toBeVisible();
  await expect.poll(() => focusIsWithin(drawer)).toBe(true);
  await expectOpaque(drawer, "the drawer");
  const drawerBox = await drawer.boundingBox();
  const scrimBox = await scrim.boundingBox();
  if (!drawerBox || !scrimBox) throw new Error("drawer or scrim has no box");
  const drawerLayoutWidth = await drawer.evaluate((element) =>
    element instanceof HTMLElement ? element.offsetWidth : null,
  );
  if (drawerLayoutWidth === null) throw new Error("drawer has no layout width");
  expect(drawerLayoutWidth).toBeLessThanOrEqual(PANEL_METRICS.drawerWidth + EPSILON);
  // Press the scrim on the side the drawer does not cover.
  const drawerOnLeft =
    drawerBox.x - scrimBox.x < scrimBox.x + scrimBox.width - (drawerBox.x + drawerBox.width);
  await page.mouse.click(
    drawerOnLeft ? scrimBox.x + scrimBox.width - 8 : scrimBox.x + 8,
    scrimBox.y + scrimBox.height / 2,
  );
  await expect(drawer).toBeHidden();
  await expect(scrim).toBeHidden();
};

/**
 * Every drawer the case offers cycles open and shut: the outline drawer from
 * the toolbar (narrow) or the rail's button (medium), and the comments drawer
 * from the comments toggle.
 */
export const expectPanelDrawers = async (page: Page, state: PanelState): Promise<void> => {
  const outlineDrawer = page.locator('[data-folio-outline-surface="drawer"]');
  if (state.outline === "drawer") {
    await expectDrawerCycle(page, {
      opener: page.getByTestId("toolbar-outline-toggle"),
      drawer: outlineDrawer,
    });
  }
  if (state.outline === "rail") {
    const rail = page.locator('[data-folio-outline-surface="rail"]');
    const tick = rail.locator(".folio-outline-tick").first();
    await tick.hover();
    await expect(tick.locator(".folio-outline-tick-label")).toBeVisible();
    await expectDrawerCycle(page, {
      opener: rail.getByTestId("folio-outline-expand"),
      drawer: outlineDrawer,
    });
  }
  if (state.comments === "drawer") {
    await expectDrawerCycle(page, {
      opener: page.getByTestId("toolbar-comments-toggle"),
      drawer: page.locator('[data-folio-comments-surface="drawer"]'),
    });
  }
};
