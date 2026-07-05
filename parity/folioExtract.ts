/**
 * Folio-side geometry extraction: loads a .docx in the playground (Playwright
 * + the existing `?file=` fixture route), walks the painted layout DOM
 * (`.layout-page` / `.layout-line`), and normalizes it into the same
 * `DocGeom` shape the Word-side extractor produces. Also captures a per-page
 * PNG screenshot for the HTML report.
 *
 * DOM-only concerns (getBoundingClientRect, closest, getComputedStyle,
 * textContent) run inside `page.evaluate` and stay pixel-valued so the
 * callback is self-contained (Playwright re-serializes it into the browser,
 * so it cannot close over module-level helpers). All arithmetic — zoom
 * normalization, px→pt conversion, font parsing, text normalization, sorting
 * — happens back in Node through `toPageGeom`, which is pure and unit-tested.
 */

import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { chromium, type Browser, type BrowserContext, type Page } from "@playwright/test";

import {
  CACHE_DIR,
  FIXTURES_DIR,
  PLAYGROUND_DEV_COMMAND,
  PLAYGROUND_URL,
  PX_TO_PT,
  REPO_ROOT,
  TMP_FIXTURE_PREFIX,
} from "./config";
import { normalizeLineText } from "./textNorm";
import type { DocGeom, LineBox, PageGeom, Region } from "./types";

export class FolioExtractError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "FolioExtractError";
  }
}

export type FolioExtraction = { geom: DocGeom; screenshotPaths: string[] };

export type FolioExtractor = {
  extract: (docxPath: string) => Promise<FolioExtraction>;
  close: () => Promise<void>;
};

const EDITOR_SELECTOR = '[data-testid="folio-editor"]';
const PAGE_SELECTOR = ".layout-page";

const VIEWPORT = { width: 1400, height: 1000 };

const SERVER_PROBE_TIMEOUT_MS = 2000;
const SERVER_START_TIMEOUT_MS = 90_000;
const SERVER_POLL_INTERVAL_MS = 500;

const EDITOR_RENDER_TIMEOUT_MS = 20_000;
const STABILITY_POLL_INTERVAL_MS = 250;
const STABILITY_MAX_MS = 15_000;
const STABILITY_SETTLE_MS = 250;

const CHROMIUM_MISSING_MARKER = "Executable doesn't exist";
const CHROMIUM_MISSING_MESSAGE =
  "Playwright chromium missing; run: bunx playwright install chromium";

/** A raw rect as returned by `getBoundingClientRect()` (visual/CSS px). */
export type RawRect = { left: number; top: number; width: number; height: number };

/** One `.layout-line` element, extracted with DOM-only reads (no math). */
export type RawLine = {
  text: string;
  rect: RawRect;
  region: Region;
  /** `getComputedStyle(...).fontFamily` of the first `.layout-run`, verbatim. */
  fontFamilyRaw?: string;
  /** `getComputedStyle(...).fontSize` of the first `.layout-run`, parsed to a px number. */
  fontSizePx?: number;
};

/** One `.layout-page` element and its lines, extracted with DOM-only reads. */
export type RawPage = {
  /** From `data-page-number`, falling back to DOM order (1-based) when absent/non-numeric. */
  pageNumber: number;
  /** Index into the page's DOM-order `.layout-page` NodeList; used to re-locate the
   * element for screenshotting without re-querying by (possibly ambiguous) page number. */
  domIndex: number;
  pageRect: RawRect;
  /** `HTMLElement.offsetWidth` / `offsetHeight` (layout px, unaffected by CSS transform zoom). */
  offsetWidth: number;
  offsetHeight: number;
  lines: RawLine[];
};

/** Ratio of layout px to visual (post-zoom) px. Guards divide-by-zero for a
 * detached/zero-size page rect by falling back to no zoom correction. */
export const computeZoomFactor = (offsetWidth: number, pageRectWidth: number): number => {
  if (!Number.isFinite(pageRectWidth) || pageRectWidth <= 0) {
    return 1;
  }
  return offsetWidth / pageRectWidth;
};

/** First font-family in a CSS `font-family` list, with surrounding quotes stripped. */
export const parseFirstFontFamily = (fontFamilyRaw: string | undefined): string | undefined => {
  if (!fontFamilyRaw) {
    return undefined;
  }
  const first = fontFamilyRaw.split(",")[0]?.trim();
  if (!first) {
    return undefined;
  }
  return first.replace(/^["']|["']$/gu, "");
};

/**
 * Pure transformer from one raw evaluate()-extracted page to the normalized
 * `PageGeom` the comparator consumes: applies the zoom-normalized px→pt
 * conversion, drops zero-size and empty-after-normalization lines, and sorts
 * the remaining lines by (yPt, xPt).
 */
export const toPageGeom = (rawPage: RawPage): PageGeom => {
  const zoomFactor = computeZoomFactor(rawPage.offsetWidth, rawPage.pageRect.width);
  const widthPt = rawPage.offsetWidth * PX_TO_PT;
  const heightPt = rawPage.offsetHeight * PX_TO_PT;

  const lines: LineBox[] = [];
  for (const rawLine of rawPage.lines) {
    if (rawLine.rect.width <= 0 || rawLine.rect.height <= 0) {
      continue;
    }
    const normText = normalizeLineText(rawLine.text);
    if (!normText) {
      continue;
    }

    const xPt = (rawLine.rect.left - rawPage.pageRect.left) * zoomFactor * PX_TO_PT;
    const yPt = (rawLine.rect.top - rawPage.pageRect.top) * zoomFactor * PX_TO_PT;
    const lineWidthPt = rawLine.rect.width * zoomFactor * PX_TO_PT;
    const lineHeightPt = rawLine.rect.height * zoomFactor * PX_TO_PT;

    const fontName = parseFirstFontFamily(rawLine.fontFamilyRaw);
    const fontSizePt = rawLine.fontSizePx !== undefined ? rawLine.fontSizePx * PX_TO_PT : undefined;

    lines.push({
      text: rawLine.text,
      normText,
      xPt,
      yPt,
      widthPt: lineWidthPt,
      heightPt: lineHeightPt,
      region: rawLine.region,
      ...(fontName !== undefined ? { fontName } : {}),
      ...(fontSizePt !== undefined ? { fontSizePt } : {}),
    });
  }

  lines.sort((a, b) => a.yPt - b.yPt || a.xPt - b.xPt);

  return { number: rawPage.pageNumber, widthPt, heightPt, lines };
};

const probeServer = async (url: string, timeoutMs: number): Promise<boolean> => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    await fetch(url, { signal: controller.signal });
    return true;
  } catch {
    return false;
  } finally {
    clearTimeout(timer);
  }
};

const waitForServerReady = async (
  url: string,
  timeoutMs: number,
  intervalMs: number,
): Promise<boolean> => {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    if (await probeServer(url, SERVER_PROBE_TIMEOUT_MS)) {
      return true;
    }
    if (Date.now() >= deadline) {
      return false;
    }
    await Bun.sleep(intervalMs);
  }
};

/**
 * `bun --filter @stll/playground dev` runs in the caller's own process group
 * (Bun.spawn doesn't detach it), so it execs into a grandchild `vite` process
 * that a plain `serverProcess.kill()` never reaches — the top-level `bun`
 * process exits but the actual dev server it spawned keeps listening and
 * leaks across runs. Walk `ps`'s pid/ppid table to find every descendant of
 * the spawned pid and signal the whole tree, descendants first so nothing
 * outlives its already-dead parent.
 */
const killProcessTree = async (rootPid: number): Promise<void> => {
  const psProcess = Bun.spawn(["ps", "-Ao", "pid,ppid"], { stdout: "pipe", stderr: "ignore" });
  const output = await new Response(psProcess.stdout).text();
  await psProcess.exited;

  const childPidsByParent = new Map<number, number[]>();
  for (const line of output.split("\n").slice(1)) {
    const [pidToken, ppidToken] = line.trim().split(/\s+/u);
    const pid = Number.parseInt(pidToken ?? "", 10);
    const ppid = Number.parseInt(ppidToken ?? "", 10);
    if (!Number.isFinite(pid) || !Number.isFinite(ppid)) {
      continue;
    }
    const siblings = childPidsByParent.get(ppid) ?? [];
    siblings.push(pid);
    childPidsByParent.set(ppid, siblings);
  }

  const treePids: number[] = [];
  const queue = [rootPid];
  for (let i = 0; i < queue.length; i++) {
    const pid = queue[i];
    if (pid === undefined) {
      continue;
    }
    treePids.push(pid);
    queue.push(...(childPidsByParent.get(pid) ?? []));
  }

  for (const pid of treePids.reverse()) {
    try {
      process.kill(pid, "SIGTERM");
    } catch {
      // Already exited — fine.
    }
  }
};

/** Poll (page count, total line count) until unchanged across 2 consecutive
 * polls (i.e. 3 identical samples in a row), capped at STABILITY_MAX_MS. Not
 * an error to time out here — the caller's final settle delay is the backstop. */
const waitForLayoutStability = async (page: Page): Promise<void> => {
  const deadline = Date.now() + STABILITY_MAX_MS;
  let lastSignature: string | null = null;
  let matchStreak = 0;

  for (;;) {
    const signature = await page.evaluate(
      () =>
        `${document.querySelectorAll(".layout-page").length}:${document.querySelectorAll(".layout-line").length}`,
    );

    if (signature === lastSignature) {
      matchStreak += 1;
      if (matchStreak >= 2) {
        return;
      }
    } else {
      matchStreak = 0;
    }
    lastSignature = signature;

    if (Date.now() >= deadline) {
      return;
    }
    await page.waitForTimeout(STABILITY_POLL_INTERVAL_MS);
  }
};

const waitForEditorLayout = async (page: Page): Promise<void> => {
  try {
    await page.waitForSelector(EDITOR_SELECTOR, { timeout: EDITOR_RENDER_TIMEOUT_MS });
  } catch {
    throw new FolioExtractError(
      `folio editor root (${EDITOR_SELECTOR}) never rendered within ${EDITOR_RENDER_TIMEOUT_MS}ms`,
    );
  }

  try {
    await page.waitForSelector(PAGE_SELECTOR, { timeout: EDITOR_RENDER_TIMEOUT_MS });
  } catch {
    throw new FolioExtractError(
      `folio never painted a ${PAGE_SELECTOR} element within ${EDITOR_RENDER_TIMEOUT_MS}ms`,
    );
  }

  await page.evaluate(() => document.fonts.ready);
  await waitForLayoutStability(page);
  await page.waitForTimeout(STABILITY_SETTLE_MS);
};

/** One `.layout-page` element's identity, listed before any scrolling so the
 * per-page extraction loop below knows what to visit and in what order. */
type PageMeta = { domIndex: number; pageNumber: number };

/** Self-contained: Playwright serializes this into the browser, so it cannot
 * reference module-level helpers. Cheap — reads only `data-page-number`, no
 * geometry — so listing it doesn't force virtualized pages to populate. */
const listPageMeta = (page: Page): Promise<PageMeta[]> =>
  page.evaluate<PageMeta[]>(() => {
    const pageEls = Array.from(document.querySelectorAll(".layout-page")) as HTMLElement[];
    const meta = pageEls.map((el, domIndex) => {
      const dataPageNumber = Number(el.dataset["pageNumber"]);
      const pageNumber = Number.isFinite(dataPageNumber) ? dataPageNumber : domIndex + 1;
      return { domIndex, pageNumber };
    });
    meta.sort((a, b) => a.pageNumber - b.pageNumber);
    return meta;
  });

/**
 * Extract one page's geometry by DOM index. Self-contained (same
 * serialization constraint as `listPageMeta`): only DOM-only reads
 * (getBoundingClientRect, closest, getComputedStyle, textContent) happen
 * here; all arithmetic lives in `toPageGeom`, back in Node.
 *
 * Folio virtualizes documents past a page-count threshold (core
 * `renderPage.ts`, `VIRTUALIZATION_THRESHOLD`): off-screen pages are
 * lightweight shells with zero `.layout-line` children until an
 * IntersectionObserver scrolls them into range, and pages far from the
 * current scroll position get depopulated again. Callers must scroll this
 * page into view (and let layout stabilize) before calling this, or it will
 * read an empty/stale shell.
 */
const extractSinglePage = (page: Page, domIndex: number): Promise<RawPage> =>
  page.evaluate<RawPage, number>((idx) => {
    const pageEls = Array.from(document.querySelectorAll(".layout-page")) as HTMLElement[];
    const el = pageEls[idx];
    if (!el) {
      throw new Error(`layout-page at domIndex ${idx} disappeared`);
    }

    const pageRect = el.getBoundingClientRect();
    const dataPageNumber = Number(el.dataset["pageNumber"]);
    const pageNumber = Number.isFinite(dataPageNumber) ? dataPageNumber : idx + 1;

    const lineEls = Array.from(el.querySelectorAll(".layout-line")) as HTMLElement[];
    const lines = lineEls.map((lineEl) => {
      // The `.layout-line` box spans the full column width regardless of where
      // the text ink actually sits (centered/right-aligned lines, table cells),
      // while the Word-side extractor reports glyph-ink bounds. Use the union
      // of the line's run and list-marker boxes as the ink rect so both sides
      // measure the same thing; fall back to the line box for lines without
      // segment children.
      const segmentEls = Array.from(
        lineEl.querySelectorAll(".layout-run, .layout-list-marker"),
      ) as HTMLElement[];
      let rect = lineEl.getBoundingClientRect();
      let inkLeft = Number.POSITIVE_INFINITY;
      let inkTop = Number.POSITIVE_INFINITY;
      let inkRight = Number.NEGATIVE_INFINITY;
      let inkBottom = Number.NEGATIVE_INFINITY;
      for (const segmentEl of segmentEls) {
        const segmentRect = segmentEl.getBoundingClientRect();
        if (segmentRect.width <= 0 || segmentRect.height <= 0) continue;
        inkLeft = Math.min(inkLeft, segmentRect.left);
        inkTop = Math.min(inkTop, segmentRect.top);
        inkRight = Math.max(inkRight, segmentRect.right);
        inkBottom = Math.max(inkBottom, segmentRect.bottom);
      }
      if (inkRight > inkLeft && inkBottom > inkTop) {
        rect = new DOMRect(inkLeft, inkTop, inkRight - inkLeft, inkBottom - inkTop);
      }
      let region: "header" | "footer" | "body" = "body";
      if (lineEl.closest(".layout-page-header")) {
        region = "header";
      } else if (lineEl.closest(".layout-page-footer")) {
        region = "footer";
      }

      const runEl = lineEl.querySelector(".layout-run");
      let fontFamilyRaw: string | undefined;
      let fontSizePx: number | undefined;
      if (runEl) {
        const computed = getComputedStyle(runEl);
        fontFamilyRaw = computed.fontFamily;
        const parsedSize = Number.parseFloat(computed.fontSize);
        if (Number.isFinite(parsedSize)) {
          fontSizePx = parsedSize;
        }
      }

      return {
        text: lineEl.textContent ?? "",
        rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
        region,
        ...(fontFamilyRaw !== undefined ? { fontFamilyRaw } : {}),
        ...(fontSizePx !== undefined ? { fontSizePx } : {}),
      };
    });

    return {
      pageNumber,
      domIndex: idx,
      pageRect: {
        left: pageRect.left,
        top: pageRect.top,
        width: pageRect.width,
        height: pageRect.height,
      },
      offsetWidth: el.offsetWidth,
      offsetHeight: el.offsetHeight,
      lines,
    };
  }, domIndex);

const stagedFixtureName = (sha256: string): string =>
  `${TMP_FIXTURE_PREFIX}${sha256.slice(0, 12)}.docx`;

export const createFolioExtractor = async (
  opts: { headless?: boolean } = {},
): Promise<FolioExtractor> => {
  const headless = opts.headless ?? true;

  let serverProcess: ReturnType<typeof Bun.spawn> | null = null;
  const serverAlreadyUp = await probeServer(PLAYGROUND_URL, SERVER_PROBE_TIMEOUT_MS);
  if (!serverAlreadyUp) {
    serverProcess = Bun.spawn(PLAYGROUND_DEV_COMMAND, {
      cwd: REPO_ROOT,
      stdin: "ignore",
      stdout: "ignore",
      stderr: "ignore",
    });
    const ready = await waitForServerReady(
      PLAYGROUND_URL,
      SERVER_START_TIMEOUT_MS,
      SERVER_POLL_INTERVAL_MS,
    );
    if (!ready) {
      await killProcessTree(serverProcess.pid);
      throw new FolioExtractError(
        `playground dev server did not become ready at ${PLAYGROUND_URL} within ${SERVER_START_TIMEOUT_MS}ms`,
      );
    }
  }
  const startedServer = serverProcess !== null;

  let browser: Browser;
  try {
    browser = await chromium.launch({ headless });
  } catch (error) {
    if (serverProcess) {
      await killProcessTree(serverProcess.pid);
    }
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes(CHROMIUM_MISSING_MARKER)) {
      throw new FolioExtractError(CHROMIUM_MISSING_MESSAGE);
    }
    throw error;
  }

  const context: BrowserContext = await browser.newContext({
    viewport: VIEWPORT,
    deviceScaleFactor: 1,
    colorScheme: "light",
  });
  const page = await context.newPage();

  const extract = async (docxPath: string): Promise<FolioExtraction> => {
    const absoluteDocxPath = path.resolve(docxPath);
    const docxBuffer = await fs.readFile(absoluteDocxPath);
    const sha256 = createHash("sha256").update(docxBuffer).digest("hex");
    const stagedName = stagedFixtureName(sha256);
    const stagedPath = path.join(FIXTURES_DIR, stagedName);

    await fs.copyFile(absoluteDocxPath, stagedPath);
    try {
      await page.goto(`${PLAYGROUND_URL}/?file=${encodeURIComponent(stagedName)}`, {
        waitUntil: "domcontentloaded",
      });
      await waitForEditorLayout(page);

      const pageMeta = await listPageMeta(page);
      if (pageMeta.length === 0) {
        throw new FolioExtractError(`folio rendered zero pages for ${absoluteDocxPath}`);
      }

      // Large documents virtualize: pages scroll into (and back out of) a
      // populated state as an IntersectionObserver crosses them, so each
      // page must be scrolled into view — forcing it to populate — right
      // before its geometry and screenshot are captured. Doing both off the
      // same scroll keeps them consistent, and small (non-virtualized)
      // documents pay only a cheap no-op scroll per page.
      const screenshotDir = path.join(CACHE_DIR, sha256, "folio-pages");
      await fs.mkdir(screenshotDir, { recursive: true });
      const rawPages: RawPage[] = [];
      const screenshotPaths: string[] = [];
      for (const { domIndex, pageNumber } of pageMeta) {
        const locator = page.locator(PAGE_SELECTOR).nth(domIndex);
        await locator.scrollIntoViewIfNeeded();
        await waitForLayoutStability(page);
        await page.waitForTimeout(STABILITY_SETTLE_MS);

        const rawPage = await extractSinglePage(page, domIndex);
        rawPages.push(rawPage);

        const screenshotPath = path.join(screenshotDir, `p${pageNumber}.png`);
        await locator.screenshot({ path: screenshotPath });
        screenshotPaths.push(screenshotPath);
      }

      const pages = rawPages.map(toPageGeom);

      const firstRawPage = rawPages[0];
      const zoomFactor = firstRawPage
        ? computeZoomFactor(firstRawPage.offsetWidth, firstRawPage.pageRect.width)
        : 1;

      const geom: DocGeom = {
        source: "folio",
        file: absoluteDocxPath,
        pages,
        meta: {
          playgroundUrl: PLAYGROUND_URL,
          stagedName,
          zoomFactor: String(zoomFactor),
          pxToPt: String(PX_TO_PT),
        },
      };

      return { geom, screenshotPaths };
    } finally {
      await fs.unlink(stagedPath).catch(() => {});
    }
  };

  const close = async (): Promise<void> => {
    await context.close();
    await browser.close();
    if (startedServer && serverProcess) {
      await killProcessTree(serverProcess.pid);
    }
  };

  return { extract, close };
};
