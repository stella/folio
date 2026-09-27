import type { Page } from "@playwright/test";

import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// A keystroke schedules a layout pass; the host then loads another revision
// (`loadDocument`) in the same task. The pass must lay out the loaded state,
// never the keystroke's, whether frames run at once or the page is hidden
// until after the adapters' timers have fired (#1142).

const MARKER = "QZXJQ";

type StaleLayoutProbe = {
  markerPainted: boolean;
  frameRequests: number;
  completedFrames: number;
  resumeFrames: () => void;
};

declare global {
  // oxlint-disable-next-line typescript/consistent-type-definitions -- Window augmentation requires an interface
  interface Window {
    __staleLayoutProbe?: StaleLayoutProbe;
  }
}

const FRAMES = {
  running: "running",
  paused: "paused",
} as const;

type Frames = (typeof FRAMES)[keyof typeof FRAMES];

/** Record whether the pages ever paint the marker; optionally hold every frame. */
const installProbe = (page: Page, frames: Frames) =>
  page.evaluate(
    ({ marker, pauseFrames }) => {
      const pages = document.querySelector(".paged-editor__pages");
      if (!pages) {
        throw new Error("no pages container");
      }
      const queued: FrameRequestCallback[] = [];
      const requestFrame = window.requestAnimationFrame.bind(window);
      const probe: StaleLayoutProbe = {
        markerPainted: false,
        frameRequests: 0,
        completedFrames: 0,
        resumeFrames: () => {
          window.requestAnimationFrame = requestFrame;
          for (const callback of queued.splice(0)) {
            callback(performance.now());
          }
        },
      };
      window.requestAnimationFrame = (callback) => {
        probe.frameRequests += 1;
        const run = (time: number) => {
          callback(time);
          probe.completedFrames += 1;
        };
        return pauseFrames ? queued.push(run) : requestFrame(run);
      };
      new MutationObserver(() => {
        const painted = [...pages.querySelectorAll(".layout-page-content")]
          .map((content) => content.textContent)
          .join("");
        if (painted.includes(marker)) {
          probe.markerPainted = true;
        }
      }).observe(pages, { childList: true, subtree: true, characterData: true });
      window.__staleLayoutProbe = probe;
    },
    { marker: MARKER, pauseFrames: frames === FRAMES.paused },
  );

const typeThenReload = async (page: Page, frames: Frames): Promise<void> => {
  await installProbe(page, frames);
  const typed = await page.evaluate(
    (marker) => window.__folioParity?.typeThenReloadDocument(marker) ?? false,
    MARKER,
  );
  expect(typed).toBe(true);
  // The edit's scheduler must request its frame before a paused tab resumes.
  await page.waitForFunction(() => (window.__staleLayoutProbe?.frameRequests ?? 0) > 0);
  await page.evaluate(() => window.__staleLayoutProbe?.resumeFrames());
  await page.waitForFunction(() => (window.__staleLayoutProbe?.completedFrames ?? 0) > 0);
  // Let the painted DOM and MutationObserver run after the scheduler callback.
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
};

for (const frames of [FRAMES.running, FRAMES.paused]) {
  forEachAdapter(
    `a keystroke's pass never paints over a loaded revision (frames ${frames})`,
    async (adapter, { page }) => {
      await openEditor(page, adapter);
      await ensureLiveView(page);

      await typeThenReload(page, frames);

      expect(
        await page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
      ).not.toContain(MARKER);
      expect(await page.evaluate(() => window.__staleLayoutProbe?.markerPainted)).toBe(false);
    },
  );
}
