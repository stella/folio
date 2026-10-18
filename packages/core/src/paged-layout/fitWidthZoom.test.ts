import { describe, expect, test } from "bun:test";

import {
  FIT_TO_WIDTH_MAX_ZOOM,
  FIT_TO_WIDTH_MIN_ZOOM,
  FIT_TO_WIDTH_PADDING_PX,
  fitWidthZoom,
} from "./fitWidthZoom";

// A4 page at 96 dpi, as laid out at zoom 1.
const PAGE_WIDTH = 794;
const PAGE_HEIGHT = 1123;
const PAGE_GAP = 24;
const VIEWPORT_PADDING = 48;
const MAX_FIT_CHANGES = 2;

type Container = {
  width: number;
  height: number;
  border: number;
  /** Classic scrollbar width; 0 models overlay scrollbars. */
  scrollbar: number;
  pages: number;
};

const scaledDocumentHeight = (pages: number, zoom: number) =>
  (pages * PAGE_HEIGHT + (pages - 1) * PAGE_GAP) * zoom + VIEWPORT_PADDING;

/** The box the browser reports: a classic scrollbar shows while the scaled document overflows. */
const boxAt = (container: Container, zoom: number) => {
  const inner = container.width - 2 * container.border;
  const overflows =
    scaledDocumentHeight(container.pages, zoom) > container.height - 2 * container.border;
  return {
    offsetWidth: container.width,
    clientWidth: inner - (overflows ? container.scrollbar : 0),
  };
};

/** Runs the hook's fit loop (each zoom change resizes the content box and refits) until it holds. */
const settle = (container: Container) => {
  let zoom = 1;
  let reservedEdge = 0;
  const zooms: number[] = [];
  for (let step = 0; step < 10; step += 1) {
    const fit = fitWidthZoom({ ...boxAt(container, zoom), pageWidth: PAGE_WIDTH, reservedEdge });
    reservedEdge = fit.reservedEdge;
    if (fit.zoom === null || Math.abs(fit.zoom - zoom) <= 0.005) {
      return { zoom, zooms, settled: true };
    }
    zoom = fit.zoom;
    zooms.push(zoom);
  }
  return { zoom, zooms, settled: false };
};

/** Heights a few pixels around the overflow threshold of every zoom the width can fit to. */
const thresholdHeights = (width: number, scrollbar: number, border: number, pages: number) => {
  const heights = new Set<number>();
  for (const gutter of [0, scrollbar]) {
    const available = width - 2 * border - gutter - FIT_TO_WIDTH_PADDING_PX;
    const zoom = Math.min(FIT_TO_WIDTH_MAX_ZOOM, Math.round((available / PAGE_WIDTH) * 100) / 100);
    const threshold = Math.ceil(scaledDocumentHeight(pages, zoom)) + 2 * border;
    for (const delta of [-3, -1, 0, 1, 3]) heights.add(threshold + delta);
  }
  return [...heights].filter((height) => height > 2 * border);
};

describe("fitWidthZoom", () => {
  test("settles within two zoom changes at every width and near-threshold height", () => {
    let containers = 0;
    for (let width = 320; width <= 1400; width += 1) {
      for (const scrollbar of [0, 15, 17]) {
        for (const border of [0, 1]) {
          for (const pages of [1, 3]) {
            for (const height of thresholdHeights(width, scrollbar, border, pages)) {
              const container = { width, height, border, scrollbar, pages };
              const result = settle(container);
              containers += 1;
              if (!result.settled || result.zooms.length > MAX_FIT_CHANGES) {
                throw new Error(
                  `fit did not settle for ${JSON.stringify(container)}: ${result.zooms.join(" -> ")}`,
                );
              }
              const box = boxAt(container, result.zoom);
              expect(result.zoom).toBeGreaterThanOrEqual(FIT_TO_WIDTH_MIN_ZOOM);
              expect(result.zoom * PAGE_WIDTH).toBeLessThanOrEqual(
                box.clientWidth - FIT_TO_WIDTH_PADDING_PX + PAGE_WIDTH * 0.005,
              );
            }
          }
        }
      }
    }
    expect(containers).toBeGreaterThan(10_000);
  });

  test("the container model reproduces the clientWidth feedback loop", () => {
    // The reported case: 427 px wide, a 17 px classic scrollbar, a height
    // between the document heights at the two adjacent fits.
    const container = { width: 427, height: 0, border: 0, scrollbar: 17, pages: 1 };
    container.height = Math.round(
      (scaledDocumentHeight(1, 0.5) + scaledDocumentHeight(1, 0.52)) / 2,
    );
    const clientWidthFit = (zoom: number) =>
      Math.round(
        ((boxAt(container, zoom).clientWidth - FIT_TO_WIDTH_PADDING_PX) / PAGE_WIDTH) * 100,
      ) / 100;
    expect(clientWidthFit(0.5)).toBe(0.52);
    expect(clientWidthFit(0.52)).toBe(0.5);

    // The reserved gutter fits once, below the threshold, and holds there.
    expect(settle(container)).toEqual({ zoom: 0.5, zooms: [0.5], settled: true });
  });

  test("keeps the widest edge it has seen", () => {
    const withBar = fitWidthZoom({
      offsetWidth: 600,
      clientWidth: 583,
      pageWidth: PAGE_WIDTH,
      reservedEdge: 0,
    });
    expect(withBar.reservedEdge).toBe(17);
    const barGone = fitWidthZoom({
      offsetWidth: 600,
      clientWidth: 600,
      pageWidth: PAGE_WIDTH,
      reservedEdge: withBar.reservedEdge,
    });
    expect(barGone).toEqual(withBar);
  });

  test("caps the fit at natural size and answers null without room", () => {
    expect(
      fitWidthZoom({ offsetWidth: 2000, clientWidth: 2000, pageWidth: PAGE_WIDTH, reservedEdge: 0 })
        .zoom,
    ).toBe(FIT_TO_WIDTH_MAX_ZOOM);
    expect(
      fitWidthZoom({ offsetWidth: 30, clientWidth: 13, pageWidth: PAGE_WIDTH, reservedEdge: 0 })
        .zoom,
    ).toBeNull();
    expect(
      fitWidthZoom({ offsetWidth: 600, clientWidth: 600, pageWidth: 0, reservedEdge: 0 }).zoom,
    ).toBeNull();
  });
});
