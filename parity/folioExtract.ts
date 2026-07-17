/**
 * Folio-side geometry extraction: loads a .docx in the playground (Playwright
 * + the existing `?file=` fixture route), walks the painted layout DOM
 * (`.layout-page` / `.layout-line`), and normalizes it into the same
 * `DocGeom` shape produced by external reference renderers. Also captures a per-page
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
  PLAYGROUND_PORT,
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

const MEANINGFUL_INK_CHARACTER = /[^\s\u00ad\u200b-\u200d\ufeff]/u;

/** Character offsets whose glyphs should contribute to a line's ink bounds. */
export const meaningfulTextRange = (text: string): { start: number; end: number } | null => {
  const first = text.search(MEANINGFUL_INK_CHARACTER);
  if (first < 0) {
    return null;
  }
  let end = text.length;
  while (end > first && !MEANINGFUL_INK_CHARACTER.test(text[end - 1] ?? "")) {
    end -= 1;
  }
  return { start: first, end };
};

export type FolioSpanInspection = {
  text: string;
  className: string;
  rect: RawRect;
  pmStart?: number;
  pmEnd?: number;
  fontFamilyRaw?: string;
  fontSizePx?: number;
  fontWeight?: string;
  fontStyle?: string;
  textTransform?: string;
};

export type FolioLineInspection = {
  index: number;
  text: string;
  rect: RawRect;
  region: Region;
  spans: FolioSpanInspection[];
};

export type FolioPageInspection = {
  pageNumber: number;
  domIndex: number;
  pageRect: RawRect;
  offsetWidth: number;
  offsetHeight: number;
  zoomFactor: number;
  lines: FolioLineInspection[];
};

export type FolioExtractOptions = {
  maxPages?: number;
};

/** A font installed alongside a local reference renderer. The harness reads
 * the file in Node and passes it to Chromium as a data URL; local paths are
 * never exposed to the playground server. */
export type LocalFontDefinition = {
  family: string;
  filePath: string;
  weight?: number | string;
  style?: string;
};

type BrowserFontDefinition = {
  family: string;
  src: string;
  weight?: number | string;
  style?: string;
};

type RoutedBrowserFont = {
  definition: BrowserFontDefinition;
  body: Buffer;
  contentType: string;
};

export type CreateFolioExtractorOptions = {
  headless?: boolean;
  reuseServer?: boolean;
  localFonts?: ReadonlyArray<LocalFontDefinition>;
};

export type FolioExtractor = {
  extract: (docxPath: string, options?: FolioExtractOptions) => Promise<FolioExtraction>;
  inspectPage: (docxPath: string, pageNumber: number) => Promise<FolioPageInspection>;
  close: () => Promise<void>;
};

const EDITOR_SELECTOR = '[data-testid="folio-editor"]';
const PAGE_SELECTOR = ".layout-page";

const VIEWPORT = { width: 1400, height: 1000 };
const SCREENSHOT_VIEWPORT_VERTICAL_CHROME_PX = 200;

// A cold Vite source transform can occupy the server before its first HTTP
// response. Short probes repeatedly abort that same warm-up work and can report
// a false startup failure even after Vite is listening.
const SERVER_PROBE_TIMEOUT_MS = 30_000;
const SERVER_OUTPUT_DRAIN_TIMEOUT_MS = 2000;
const SERVER_REUSE_PROBE_TIMEOUT_MS = 15_000;
const SERVER_START_TIMEOUT_MS = 180_000;
const SERVER_POLL_INTERVAL_MS = 500;
const SERVER_LOG_TAIL_CHARS = 4000;

const PLAYGROUND_NAVIGATION_TIMEOUT_MS = 300_000;
const EDITOR_RENDER_TIMEOUT_MS = 90_000;
const STABILITY_POLL_INTERVAL_MS = 250;
const STABILITY_MAX_MS = 15_000;
const STABILITY_SETTLE_MS = 250;

const CHROMIUM_MISSING_MARKER = "Executable doesn't exist";
const CHROMIUM_MISSING_MESSAGE =
  "Playwright chromium missing; run: bunx playwright install chromium";

const FONT_MIME_BY_EXTENSION = {
  ".otf": "font/otf",
  ".ttf": "font/ttf",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
} as const;

export const localFontContentType = (filePath: string): string => {
  const extension = path.extname(filePath).toLowerCase();
  const mime = FONT_MIME_BY_EXTENSION[extension as keyof typeof FONT_MIME_BY_EXTENSION];
  if (!mime) {
    throw new FolioExtractError(`unsupported local font format: ${extension || "(none)"}`);
  }
  return mime;
};

export const localFontRouteUrl = (index: number, filePath: string): string =>
  `${PLAYGROUND_URL}/__folio-parity-font/${index}${path.extname(filePath).toLowerCase()}`;

const loadBrowserFonts = async (
  fonts: ReadonlyArray<LocalFontDefinition> | undefined,
): Promise<RoutedBrowserFont[]> =>
  await Promise.all(
    (fonts ?? []).map(async ({ family, filePath, weight, style }, index) => {
      const definition: BrowserFontDefinition = {
        family,
        src: localFontRouteUrl(index, filePath),
      };
      if (weight !== undefined) {
        definition.weight = weight;
      }
      if (style !== undefined) {
        definition.style = style;
      }
      return {
        definition,
        body: await fs.readFile(filePath),
        contentType: localFontContentType(filePath),
      };
    }),
  );

type OutputTail = { read: () => string; done: Promise<void> };

const captureOutputTail = (stream: ReadableStream<Uint8Array> | null | undefined): OutputTail => {
  const decoder = new TextDecoder();
  let output = "";

  if (!stream) {
    return { read: () => "", done: Promise.resolve() };
  }

  const done = (async () => {
    try {
      for await (const chunk of stream) {
        output = `${output}${decoder.decode(chunk, { stream: true })}`.slice(
          -SERVER_LOG_TAIL_CHARS,
        );
      }
      output = `${output}${decoder.decode()}`.slice(-SERVER_LOG_TAIL_CHARS);
    } catch (error) {
      output = `${output}\n[log capture failed: ${String(error)}]`.slice(-SERVER_LOG_TAIL_CHARS);
    }
  })();

  return { read: () => output.trim(), done };
};

export const formatServerStartFailure = (
  url: string,
  timeoutMs: number,
  exitCode: number | undefined,
  output: string,
): string => {
  const reason =
    exitCode === undefined
      ? `did not become ready at ${url} within ${timeoutMs}ms`
      : `exited with code ${exitCode} before becoming ready at ${url}`;
  return output.length === 0
    ? `playground dev server ${reason}`
    : `playground dev server ${reason}\nPlayground output:\n${output}`;
};

export const formatNavigationFailure = (url: string, error: unknown, output: string): string => {
  const message = error instanceof Error ? error.message : String(error);
  return output.length === 0
    ? `playground navigation failed for ${url}: ${message}`
    : `playground navigation failed for ${url}: ${message}\nPlayground output:\n${output}`;
};

/** A raw rect as returned by `getBoundingClientRect()` (visual/CSS px). */
export type RawRect = { left: number; top: number; width: number; height: number };

/** One ink segment extracted from a `.layout-line` with DOM-only reads (no math). */
export type RawLine = {
  text: string;
  rect: RawRect;
  /** Viewport-relative CSS-pixel position of the originating line's baseline. */
  baselineTop?: number;
  region: Region;
  /** True when the line's ink box falls fully outside an overflow-clipping ancestor. */
  fullyClipped?: boolean;
  /** CSS clipping ancestors, captured in visual pixels for visibility filtering. */
  clippingAncestors?: RawClippingAncestor[];
  /** First actually available family from the computed CSS stack of the first
   * `.layout-run`; falls back to the computed stack when canvas probing is unavailable. */
  fontFamilyRaw?: string;
  /** `getComputedStyle(...).fontSize` of the first `.layout-run`, parsed to a px number. */
  fontSizePx?: number;
  /** Stable page-local table-cell identity, when the line is inside a cell. */
  visualGroup?: string;
  /** Stable page-local identity of the originating `.layout-line`. */
  logicalLineGroup?: string;
};

export type RawClippingAncestor = {
  rect: RawRect;
  offsetWidth: number;
  offsetHeight: number;
  clipsX: boolean;
  clipsY: boolean;
  clipPath?: string;
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

const MIN_BASELINE_OFFSET_RATIO = -0.5;
const MAX_BASELINE_OFFSET_RATIO = 1.5;

/**
 * Baseline probes are zero-width inline boxes. On a visually full line, a
 * browser can wrap the probe onto the next row even though the document text
 * itself did not wrap. Keep only probe positions near the extracted ink box;
 * the comparator can fall back to ink-top geometry for rejected probes.
 */
export const isPlausibleBaseline = (baselineTop: number, rect: RawRect): boolean => {
  if (!Number.isFinite(baselineTop) || rect.height <= 0) {
    return false;
  }
  const offsetRatio = (baselineTop - rect.top) / rect.height;
  return offsetRatio >= MIN_BASELINE_OFFSET_RATIO && offsetRatio <= MAX_BASELINE_OFFSET_RATIO;
};

type InsetClip = { top: number; right: number; bottom: number; left: number };

const cssInsetLength = (token: string, dimension: number): number | null => {
  if (token === "0") return 0;
  if (token.endsWith("px")) {
    const value = Number.parseFloat(token);
    return Number.isFinite(value) ? value : null;
  }
  if (token.endsWith("%")) {
    const value = Number.parseFloat(token);
    return Number.isFinite(value) ? (value / 100) * dimension : null;
  }
  return null;
};

/** Parse the computed form of a CSS `inset()` clip path. Unsupported units
 * return null so the visibility filter remains conservative. */
export const parseInsetClipPath = (
  clipPath: string | undefined,
  width: number,
  height: number,
): InsetClip | null => {
  const match = clipPath?.match(/^inset\(([^)]*)\)$/u);
  const value = match?.[1]?.split(/\s+round\s+/u)[0]?.trim();
  if (!value) return null;

  const tokens = value.split(/\s+/u);
  if (tokens.length < 1 || tokens.length > 4) return null;
  const [topToken, rightToken = topToken, bottomToken = topToken, leftToken = rightToken] = tokens;
  if (!topToken || !rightToken || !bottomToken || !leftToken) return null;

  const top = cssInsetLength(topToken, height);
  const right = cssInsetLength(rightToken, width);
  const bottom = cssInsetLength(bottomToken, height);
  const left = cssInsetLength(leftToken, width);
  if (top === null || right === null || bottom === null || left === null) return null;
  return { top, right, bottom, left };
};

/** Whether a visual rect has no remaining area after all supported CSS clips
 * are intersected. Clip-path lengths are layout pixels, so they are scaled to
 * the post-transform visual rect before comparison. */
export const isFullyClippedByAncestors = (
  rect: RawRect,
  ancestors: RawClippingAncestor[] | undefined,
): boolean => {
  let left = rect.left;
  let top = rect.top;
  let right = rect.left + rect.width;
  let bottom = rect.top + rect.height;

  for (const ancestor of ancestors ?? []) {
    const ancestorRight = ancestor.rect.left + ancestor.rect.width;
    const ancestorBottom = ancestor.rect.top + ancestor.rect.height;
    if (ancestor.clipsX) {
      left = Math.max(left, ancestor.rect.left);
      right = Math.min(right, ancestorRight);
    }
    if (ancestor.clipsY) {
      top = Math.max(top, ancestor.rect.top);
      bottom = Math.min(bottom, ancestorBottom);
    }

    const inset = parseInsetClipPath(
      ancestor.clipPath,
      ancestor.offsetWidth,
      ancestor.offsetHeight,
    );
    if (inset) {
      const scaleX = ancestor.offsetWidth > 0 ? ancestor.rect.width / ancestor.offsetWidth : 1;
      const scaleY = ancestor.offsetHeight > 0 ? ancestor.rect.height / ancestor.offsetHeight : 1;
      left = Math.max(left, ancestor.rect.left + inset.left * scaleX);
      right = Math.min(right, ancestorRight - inset.right * scaleX);
      top = Math.max(top, ancestor.rect.top + inset.top * scaleY);
      bottom = Math.min(bottom, ancestorBottom - inset.bottom * scaleY);
    }

    if (right <= left || bottom <= top) return true;
  }
  return false;
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

const CSS_FONT_FAMILY_TOKEN_PATTERN = `"[^"]*"|'[^']*'|[^,]+`;

/** Split a computed CSS family stack without treating commas inside quoted
 * family names as separators. */
export const parseCssFontFamilies = (fontFamilyRaw: string): string[] =>
  (fontFamilyRaw.match(new RegExp(CSS_FONT_FAMILY_TOKEN_PATTERN, "gu")) ?? [])
    .map((family) => family.trim().replace(/^['"]|['"]$/gu, ""))
    .filter((family) => family.length > 0);

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
    if (
      rawLine.fullyClipped ||
      isFullyClippedByAncestors(rawLine.rect, rawLine.clippingAncestors) ||
      rawLine.rect.width <= 0 ||
      rawLine.rect.height <= 0
    ) {
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
    const baselinePt =
      rawLine.baselineTop === undefined || !isPlausibleBaseline(rawLine.baselineTop, rawLine.rect)
        ? undefined
        : (rawLine.baselineTop - rawPage.pageRect.top) * zoomFactor * PX_TO_PT;

    const fontName = parseFirstFontFamily(rawLine.fontFamilyRaw);
    const fontSizePt = rawLine.fontSizePx !== undefined ? rawLine.fontSizePx * PX_TO_PT : undefined;

    lines.push({
      text: rawLine.text,
      normText,
      xPt,
      yPt,
      ...(baselinePt !== undefined ? { baselinePt } : {}),
      widthPt: lineWidthPt,
      heightPt: lineHeightPt,
      region: rawLine.region,
      ...(rawLine.visualGroup !== undefined ? { visualGroup: rawLine.visualGroup } : {}),
      ...(rawLine.logicalLineGroup !== undefined
        ? { logicalLineGroup: rawLine.logicalLineGroup }
        : {}),
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
    const res = await fetch(url, { signal: controller.signal });
    await res.body?.cancel();
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
  signal?: AbortSignal,
): Promise<boolean> => {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    if (signal?.aborted) {
      return false;
    }
    if (await probeServer(url, SERVER_PROBE_TIMEOUT_MS)) {
      return true;
    }
    if (signal?.aborted || Date.now() >= deadline) {
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

/**
 * Kill whatever is LISTENING on `port` — the reliable, port-scoped way to
 * clear a leftover playground server before starting a fresh one.
 *
 * This is deliberately NOT `pkill -f vite`: that pattern-kill hits EVERY Vite
 * process on the machine, so it silently murders the dev servers of other
 * git worktrees / other sessions (observed clobbering an unrelated worktree's
 * server). `lsof -tiTCP:<port>` targets exactly the one server on our own
 * per-worktree port and nothing else.
 */
const killByPort = async (port: number): Promise<void> => {
  const lsof = Bun.spawn(["lsof", `-tiTCP:${port}`, "-sTCP:LISTEN"], {
    stdout: "pipe",
    stderr: "ignore",
  });
  const out = (await new Response(lsof.stdout).text()).trim();
  await lsof.exited;
  if (out.length === 0) return;
  for (const token of out.split("\n")) {
    const pid = Number.parseInt(token.trim(), 10);
    if (!Number.isFinite(pid)) continue;
    try {
      process.kill(pid, "SIGTERM");
    } catch {
      // Already gone — fine.
    }
  }
  // Give the OS a moment to release the port before we rebind it.
  await Bun.sleep(SERVER_POLL_INTERVAL_MS);
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

export const CLEAN_SCREENSHOT_CSS = `
      .pg-collab-header,
      .pg-controls,
      [data-folio-toolbar="true"],
      [data-testid="playground-controls"],
      [class*="toolbar" i],
      [class*="ruler" i],
      [style*="position: fixed"],
      [style*="position: sticky"] {
        visibility: hidden !important;
      }

      .layout-page,
      .layout-page * {
        visibility: visible !important;
      }
    `;

const installCleanScreenshotStyle = async (page: Page): Promise<void> => {
  await page.addStyleTag({
    content: CLEAN_SCREENSHOT_CSS,
  });
};

export const screenshotViewportHeight = (pageHeights: number[], currentHeight: number): number => {
  const tallestPage = Math.max(0, ...pageHeights.filter(Number.isFinite));
  return Math.max(currentHeight, Math.ceil(tallestPage + SCREENSHOT_VIEWPORT_VERTICAL_CHROME_PX));
};

const fitViewportToPages = async (page: Page): Promise<void> => {
  const viewport = page.viewportSize();
  if (!viewport) {
    return;
  }
  const pageHeights = await page
    .locator(PAGE_SELECTOR)
    .evaluateAll((pageEls) => pageEls.map((pageEl) => (pageEl as HTMLElement).offsetHeight));
  const height = screenshotViewportHeight(pageHeights, viewport.height);
  if (height === viewport.height) {
    return;
  }
  await page.setViewportSize({ width: viewport.width, height });
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
export const extractSinglePage = (page: Page, domIndex: number): Promise<RawPage> =>
  page.evaluate<
    RawPage,
    { domIndex: number; meaningfulInkPattern: string; fontFamilyTokenPattern: string }
  >(
    (input) => {
      const { domIndex: idx, meaningfulInkPattern, fontFamilyTokenPattern } = input;
      const meaningfulInkCharacter = new RegExp(meaningfulInkPattern, "u");
      const pageEls = Array.from(document.querySelectorAll(".layout-page")) as HTMLElement[];
      const el = pageEls[idx];
      if (!el) {
        throw new Error(`layout-page at domIndex ${idx} disappeared`);
      }

      const pageRect = el.getBoundingClientRect();
      const dataPageNumber = Number(el.dataset["pageNumber"]);
      const pageNumber = Number.isFinite(dataPageNumber) ? dataPageNumber : idx + 1;

      const tableCells = Array.from(el.querySelectorAll(".layout-table-cell"));
      const lineEls = Array.from(el.querySelectorAll(".layout-line")) as HTMLElement[];
      const resolvedFontCache = new Map<string, string>();
      const clippingValues = new Set(["auto", "clip", "hidden", "scroll"]);
      const ancestorClipCache = new Map<
        HTMLElement,
        {
          clipsX: boolean;
          clipsY: boolean;
          clipPath?: string;
          rect?: DOMRect;
          offsetWidth: number;
          offsetHeight: number;
        }
      >();
      const canvasContext = document.createElement("canvas").getContext("2d");
      const fontProbeText = "mmmmmmmmmmlliWW0123456789";
      const genericFamilies = new Set([
        "serif",
        "sans-serif",
        "monospace",
        "cursive",
        "fantasy",
        "system-ui",
      ]);
      const resolveFontFamily = (computed: CSSStyleDeclaration): string => {
        const cached = resolvedFontCache.get(computed.fontFamily);
        if (cached !== undefined) return cached;

        const families =
          computed.fontFamily
            .match(new RegExp(fontFamilyTokenPattern, "gu"))
            ?.map((family) => family.trim().replace(/^['"]|['"]$/gu, ""))
            .filter((family) => family.length > 0) ?? [];
        if (!canvasContext) return families[0] ?? computed.fontFamily;

        const fontSize = computed.fontSize || "16px";
        const baselines = ["monospace", "serif", "sans-serif"];
        for (const family of families) {
          if (genericFamilies.has(family.toLocaleLowerCase())) {
            resolvedFontCache.set(computed.fontFamily, family);
            return family;
          }
          const escapedFamily = family.replace(/\\/gu, "\\\\").replace(/"/gu, '\\"');
          const isAvailable = baselines.some((baseline) => {
            canvasContext.font = `${fontSize} ${baseline}`;
            const baselineWidth = canvasContext.measureText(fontProbeText).width;
            canvasContext.font = `${fontSize} "${escapedFamily}", ${baseline}`;
            const candidateWidth = canvasContext.measureText(fontProbeText).width;
            return Math.abs(candidateWidth - baselineWidth) > 0.01;
          });
          if (isAvailable) {
            resolvedFontCache.set(computed.fontFamily, family);
            return family;
          }
        }

        const fallback = families.at(-1) ?? computed.fontFamily;
        resolvedFontCache.set(computed.fontFamily, fallback);
        return fallback;
      };
      const lines = lineEls.flatMap((lineEl, lineIndex) => {
        const baselineProbe = document.createElement("span");
        baselineProbe.setAttribute("aria-hidden", "true");
        baselineProbe.style.display = "inline-block";
        baselineProbe.style.width = "0";
        baselineProbe.style.height = "0";
        baselineProbe.style.margin = "0";
        baselineProbe.style.padding = "0";
        baselineProbe.style.border = "0";
        baselineProbe.style.verticalAlign = "baseline";
        lineEl.append(baselineProbe);
        const baselineTop = baselineProbe.getBoundingClientRect().top;
        baselineProbe.remove();

        const segmentInkRect = (segmentEl: HTMLElement): DOMRect | null => {
          if (
            segmentEl.matches("img, svg, canvas, video") ||
            segmentEl.querySelector("img, svg, canvas, video")
          ) {
            const rect = segmentEl.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0 ? rect : null;
          }

          const walker = document.createTreeWalker(segmentEl, NodeFilter.SHOW_TEXT);
          let firstNode: Text | null = null;
          let firstOffset = 0;
          let lastNode: Text | null = null;
          let lastOffset = 0;
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            const textNode = node as Text;
            const value = textNode.data;
            const start = value.search(meaningfulInkCharacter);
            if (start < 0) {
              continue;
            }
            let end = value.length;
            while (end > start && !meaningfulInkCharacter.test(value[end - 1] ?? "")) {
              end -= 1;
            }
            firstNode ??= textNode;
            if (firstNode === textNode) {
              firstOffset = start;
            }
            lastNode = textNode;
            lastOffset = end;
          }

          if (firstNode && lastNode) {
            const range = document.createRange();
            range.setStart(firstNode, firstOffset);
            range.setEnd(lastNode, lastOffset);
            const rect = range.getBoundingClientRect();
            range.detach();
            if (rect.width > 0 && rect.height > 0) {
              return rect;
            }
          }

          return null;
        };
        const rectFor = (segmentEls: HTMLElement[]): DOMRect | null => {
          let inkLeft = Number.POSITIVE_INFINITY;
          let inkTop = Number.POSITIVE_INFINITY;
          let inkRight = Number.NEGATIVE_INFINITY;
          let inkBottom = Number.NEGATIVE_INFINITY;
          for (const segmentEl of segmentEls) {
            const segmentRect = segmentInkRect(segmentEl);
            if (!segmentRect) continue;
            inkLeft = Math.min(inkLeft, segmentRect.left);
            inkTop = Math.min(inkTop, segmentRect.top);
            inkRight = Math.max(inkRight, segmentRect.right);
            inkBottom = Math.max(inkBottom, segmentRect.bottom);
          }
          return inkRight > inkLeft && inkBottom > inkTop
            ? new DOMRect(inkLeft, inkTop, inkRight - inkLeft, inkBottom - inkTop)
            : null;
        };
        const region: Region = (() => {
          if (lineEl.closest(".layout-page-header")) {
            return "header";
          }
          if (lineEl.closest(".layout-page-footer")) {
            return "footer";
          }
          return "body";
        })();
        const tableCell = lineEl.closest(".layout-table-cell");
        const visualGroup = tableCell ? `table-cell:${tableCells.indexOf(tableCell)}` : undefined;
        const logicalLineGroup = `layout-line:${lineIndex}`;

        const fontFrom = (sourceEl: HTMLElement | null) => {
          if (!sourceEl) return {};
          const computed = getComputedStyle(sourceEl);
          const parsedSize = Number.parseFloat(computed.fontSize);
          return {
            fontFamilyRaw: resolveFontFamily(computed),
            ...(Number.isFinite(parsedSize) ? { fontSizePx: parsedSize } : {}),
          };
        };
        const textFor = (sourceEl: HTMLElement): string => {
          const text = sourceEl.textContent ?? "";
          const computed = getComputedStyle(sourceEl);
          if (
            computed.textTransform === "uppercase" ||
            computed.fontVariant.includes("small-caps")
          ) {
            return text.toLocaleUpperCase();
          }
          return text;
        };

        const clippingAncestorsFor = (sourceEl: HTMLElement | null) => {
          const clippingAncestors = [];
          let ancestor: HTMLElement | null = sourceEl ?? lineEl;
          while (ancestor && el.contains(ancestor)) {
            let clipInfo = ancestorClipCache.get(ancestor);
            if (!clipInfo) {
              const computed = getComputedStyle(ancestor);
              const clipsX = clippingValues.has(computed.overflowX);
              const clipsY = clippingValues.has(computed.overflowY);
              const clipPath = computed.clipPath === "none" ? undefined : computed.clipPath;
              clipInfo = {
                clipsX,
                clipsY,
                ...(clipPath !== undefined ? { clipPath } : {}),
                ...(clipsX || clipsY || clipPath !== undefined
                  ? { rect: ancestor.getBoundingClientRect() }
                  : {}),
                offsetWidth: ancestor.offsetWidth,
                offsetHeight: ancestor.offsetHeight,
              };
              ancestorClipCache.set(ancestor, clipInfo);
            }
            const {
              clipsX,
              clipsY,
              clipPath,
              rect: ancestorRect,
              offsetWidth,
              offsetHeight,
            } = clipInfo;
            if (clipsX || clipsY || clipPath !== undefined) {
              if (!ancestorRect) {
                throw new Error("clipping ancestor is missing cached geometry");
              }
              clippingAncestors.push({
                rect: {
                  left: ancestorRect.left,
                  top: ancestorRect.top,
                  width: ancestorRect.width,
                  height: ancestorRect.height,
                },
                clipsX,
                clipsY,
                ...(clipPath !== undefined ? { clipPath } : {}),
                offsetWidth,
                offsetHeight,
              });
            }
            if (ancestor === el) {
              break;
            }
            ancestor = ancestor.parentElement;
          }
          return clippingAncestors;
        };

        const toRawLine = (text: string, rect: DOMRect, sourceEl: HTMLElement | null) => ({
          text,
          rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
          ...(Number.isFinite(baselineTop) ? { baselineTop } : {}),
          region,
          clippingAncestors: clippingAncestorsFor(sourceEl),
          ...(visualGroup !== undefined ? { visualGroup } : {}),
          logicalLineGroup,
          ...fontFrom(sourceEl),
        });

        const markerEls = Array.from(
          lineEl.querySelectorAll(".layout-list-marker"),
        ) as HTMLElement[];
        const runEls = Array.from(lineEl.querySelectorAll(".layout-run")) as HTMLElement[];
        if (markerEls.length > 0 && runEls.length > 0) {
          const markerRect = rectFor(markerEls);
          const runsRect = rectFor(runEls);
          const splitLines = [];
          if (markerRect) {
            splitLines.push(
              toRawLine(
                markerEls.map((markerEl) => textFor(markerEl)).join(""),
                markerRect,
                markerEls[0] ?? null,
              ),
            );
          }
          if (runsRect) {
            splitLines.push(
              toRawLine(
                runEls.map((runEl) => textFor(runEl)).join(""),
                runsRect,
                runEls[0] ?? null,
              ),
            );
          }
          if (splitLines.length > 0) {
            return splitLines;
          }
        }

        const visualRunEls = Array.from(lineEl.children).filter(
          (child): child is HTMLElement =>
            child instanceof HTMLElement && child.classList.contains("layout-run"),
        );
        if (visualRunEls.some((runEl) => runEl.classList.contains("layout-run-tab"))) {
          const segments: HTMLElement[][] = [];
          let currentSegment: HTMLElement[] = [];
          const flushSegment = () => {
            if (currentSegment.length === 0) {
              return;
            }
            segments.push(currentSegment);
            currentSegment = [];
          };

          for (const runEl of visualRunEls) {
            if (!runEl.classList.contains("layout-run-tab")) {
              currentSegment.push(runEl);
              continue;
            }

            const tabText = (runEl.textContent ?? "").replace(/[­​-‍﻿\s]/gu, "");
            if (tabText.length > 0) {
              currentSegment.push(runEl);
              continue;
            }

            flushSegment();
          }
          flushSegment();

          if (segments.length > 0) {
            const splitLines = [];
            for (const segment of segments) {
              const segmentRect = rectFor(segment);
              if (!segmentRect) {
                continue;
              }
              const sourceEl =
                segment.find((segmentEl) => !segmentEl.classList.contains("layout-run-tab")) ??
                segment[0] ??
                null;
              splitLines.push(
                toRawLine(
                  segment.map((segmentEl) => textFor(segmentEl)).join(""),
                  segmentRect,
                  sourceEl,
                ),
              );
            }
            if (splitLines.length > 0) {
              return splitLines;
            }
          }
        }

        // The `.layout-line` box spans the full column width regardless of where
        // the text ink actually sits (centered/right-aligned lines, table cells),
        // while PDF-based reference extractors report glyph-ink bounds. Use the union
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
          const segmentRect = segmentInkRect(segmentEl);
          if (!segmentRect) continue;
          inkLeft = Math.min(inkLeft, segmentRect.left);
          inkTop = Math.min(inkTop, segmentRect.top);
          inkRight = Math.max(inkRight, segmentRect.right);
          inkBottom = Math.max(inkBottom, segmentRect.bottom);
        }
        if (inkRight > inkLeft && inkBottom > inkTop) {
          rect = new DOMRect(inkLeft, inkTop, inkRight - inkLeft, inkBottom - inkTop);
        }

        const runEl = lineEl.querySelector<HTMLElement>(".layout-run");

        const text =
          segmentEls.length > 0
            ? segmentEls.map((segmentEl) => textFor(segmentEl)).join("")
            : textFor(lineEl);
        return [toRawLine(text, rect, runEl)];
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
    },
    {
      domIndex,
      meaningfulInkPattern: MEANINGFUL_INK_CHARACTER.source,
      fontFamilyTokenPattern: CSS_FONT_FAMILY_TOKEN_PATTERN,
    },
  );

const inspectSinglePage = (page: Page, domIndex: number): Promise<FolioPageInspection> =>
  page.evaluate<FolioPageInspection, number>((idx) => {
    const pageEls = Array.from(document.querySelectorAll(".layout-page")) as HTMLElement[];
    const el = pageEls[idx];
    if (!el) {
      throw new Error(`layout-page at domIndex ${idx} disappeared`);
    }

    const toRawRect = (rect: DOMRect) => ({
      left: rect.left,
      top: rect.top,
      width: rect.width,
      height: rect.height,
    });
    const numberAttr = Number(el.dataset["pageNumber"]);
    const pageNumber = Number.isFinite(numberAttr) ? numberAttr : idx + 1;
    const pageRect = el.getBoundingClientRect();
    const zoomFactor =
      Number.isFinite(pageRect.width) && pageRect.width > 0 ? el.offsetWidth / pageRect.width : 1;

    const lineEls = Array.from(el.querySelectorAll(".layout-line")) as HTMLElement[];
    const lines = lineEls.map((lineEl, lineIndex) => {
      const region: Region = (() => {
        if (lineEl.closest(".layout-page-header")) {
          return "header";
        }
        if (lineEl.closest(".layout-page-footer")) {
          return "footer";
        }
        return "body";
      })();

      const spanEls = Array.from(
        lineEl.querySelectorAll(".layout-run, .layout-list-marker"),
      ) as HTMLElement[];
      const spans = spanEls.map((spanEl) => {
        const computed = getComputedStyle(spanEl);
        const pmStart = Number(spanEl.dataset["pmStart"]);
        const pmEnd = Number(spanEl.dataset["pmEnd"]);
        const fontSizePx = Number.parseFloat(computed.fontSize);
        return {
          text: spanEl.textContent ?? "",
          className: spanEl.className,
          rect: toRawRect(spanEl.getBoundingClientRect()),
          ...(Number.isFinite(pmStart) ? { pmStart } : {}),
          ...(Number.isFinite(pmEnd) ? { pmEnd } : {}),
          fontFamilyRaw: computed.fontFamily,
          ...(Number.isFinite(fontSizePx) ? { fontSizePx } : {}),
          fontWeight: computed.fontWeight,
          fontStyle: computed.fontStyle,
          textTransform: computed.textTransform,
        };
      });

      return {
        index: lineIndex + 1,
        text:
          spanEls.length > 0
            ? spanEls.map((spanEl) => spanEl.textContent ?? "").join("")
            : (lineEl.textContent ?? ""),
        rect: toRawRect(lineEl.getBoundingClientRect()),
        region,
        spans,
      };
    });

    return {
      pageNumber,
      domIndex: idx,
      pageRect: toRawRect(pageRect),
      offsetWidth: el.offsetWidth,
      offsetHeight: el.offsetHeight,
      zoomFactor,
      lines,
    };
  }, domIndex);

const stagedFixtureName = (sha256: string): string =>
  `${TMP_FIXTURE_PREFIX}${sha256.slice(0, 12)}.docx`;

export const createFolioExtractor = async (
  opts: CreateFolioExtractorOptions = {},
): Promise<FolioExtractor> => {
  const headless = opts.headless ?? true;
  // Default to a FRESH server so the feedback loop always reflects this
  // worktree's current source. Reusing a server risks measuring stale code:
  // under Vite `strictPort`, a leftover server on the port means a new `dev`
  // fails to bind and the OLD one keeps serving — which silently produced
  // wrong geometry until diagnosed. `reuseServer: true` opts into reuse for
  // speed when you know the running server is current.
  const reuseServer = opts.reuseServer ?? false;

  let serverProcess: ReturnType<typeof Bun.spawn> | null = null;
  let serverStdoutTail: OutputTail | null = null;
  let serverStderrTail: OutputTail | null = null;
  const serverAlreadyUp =
    reuseServer && (await probeServer(PLAYGROUND_URL, SERVER_REUSE_PROBE_TIMEOUT_MS));
  if (!serverAlreadyUp) {
    // Clear any leftover server on OUR per-worktree port (see killByPort:
    // port-scoped, so other worktrees' servers are untouched), then start
    // fresh, telling the playground which port to bind via env.
    await killByPort(PLAYGROUND_PORT);
    serverProcess = Bun.spawn(PLAYGROUND_DEV_COMMAND, {
      cwd: REPO_ROOT,
      env: { ...process.env, FOLIO_PLAYGROUND_PORT: String(PLAYGROUND_PORT) },
      stdin: "ignore",
      stdout: "pipe",
      stderr: "pipe",
    });
    serverStdoutTail = captureOutputTail(
      typeof serverProcess.stdout === "number" ? undefined : serverProcess.stdout,
    );
    serverStderrTail = captureOutputTail(
      typeof serverProcess.stderr === "number" ? undefined : serverProcess.stderr,
    );
    const startupProbe = new AbortController();
    const startup = await Promise.race([
      waitForServerReady(
        PLAYGROUND_URL,
        SERVER_START_TIMEOUT_MS,
        SERVER_POLL_INTERVAL_MS,
        startupProbe.signal,
      ).then((ready) => ({ type: "probe" as const, ready })),
      serverProcess.exited.then((exitCode) => ({ type: "exit" as const, exitCode })),
    ]);
    startupProbe.abort();
    if (startup.type === "exit" || !startup.ready) {
      await killProcessTree(serverProcess.pid);
      await killByPort(PLAYGROUND_PORT);
      await Promise.race([
        Promise.all([serverStdoutTail.done, serverStderrTail.done]),
        Bun.sleep(SERVER_OUTPUT_DRAIN_TIMEOUT_MS),
      ]);
      throw new FolioExtractError(
        formatServerStartFailure(
          PLAYGROUND_URL,
          SERVER_START_TIMEOUT_MS,
          startup.type === "exit" ? startup.exitCode : undefined,
          [serverStdoutTail.read(), serverStderrTail.read()].filter(Boolean).join("\n"),
        ),
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
  const routedFonts = await loadBrowserFonts(opts.localFonts);
  if (routedFonts.length > 0) {
    for (const { definition, body, contentType } of routedFonts) {
      // oxlint-disable-next-line no-await-in-loop -- routes must be installed before the first navigation
      await page.route(definition.src, async (route) => {
        await route.fulfill({ body, contentType });
      });
    }
    await page.addInitScript(
      (fonts) => {
        Reflect.set(globalThis, "__folioParityFonts", fonts);
      },
      routedFonts.map(({ definition }) => definition),
    );
  }

  const navigateToDocument = async (stagedName: string): Promise<void> => {
    const documentUrl = `${PLAYGROUND_URL}/?file=${encodeURIComponent(stagedName)}`;
    try {
      await page.goto(documentUrl, {
        waitUntil: "domcontentloaded",
        timeout: PLAYGROUND_NAVIGATION_TIMEOUT_MS,
      });
    } catch (error) {
      const output = [serverStdoutTail?.read(), serverStderrTail?.read()]
        .filter(Boolean)
        .join("\n");
      throw new FolioExtractError(formatNavigationFailure(documentUrl, error, output));
    }
  };

  const extract = async (
    docxPath: string,
    options: FolioExtractOptions = {},
  ): Promise<FolioExtraction> => {
    const absoluteDocxPath = path.resolve(docxPath);
    const docxBuffer = await fs.readFile(absoluteDocxPath);
    const sha256 = createHash("sha256").update(docxBuffer).digest("hex");
    const stagedName = stagedFixtureName(sha256);
    const stagedPath = path.join(FIXTURES_DIR, stagedName);

    await fs.copyFile(absoluteDocxPath, stagedPath);
    try {
      await navigateToDocument(stagedName);
      await waitForEditorLayout(page);

      const pageMeta = await listPageMeta(page);
      if (pageMeta.length === 0) {
        throw new FolioExtractError(`folio rendered zero pages for ${absoluteDocxPath}`);
      }
      await fitViewportToPages(page);
      await installCleanScreenshotStyle(page);
      const pagesToExtract =
        options.maxPages === undefined ? pageMeta : pageMeta.slice(0, options.maxPages);

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
      for (const { domIndex, pageNumber } of pagesToExtract) {
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
          localFontFaces: String(routedFonts.length),
        },
      };

      return { geom, screenshotPaths };
    } finally {
      await fs.unlink(stagedPath).catch(() => {});
    }
  };

  const inspectPage = async (
    docxPath: string,
    pageNumber: number,
  ): Promise<FolioPageInspection> => {
    const absoluteDocxPath = path.resolve(docxPath);
    const docxBuffer = await fs.readFile(absoluteDocxPath);
    const sha256 = createHash("sha256").update(docxBuffer).digest("hex");
    const stagedName = stagedFixtureName(sha256);
    const stagedPath = path.join(FIXTURES_DIR, stagedName);

    await fs.copyFile(absoluteDocxPath, stagedPath);
    try {
      await navigateToDocument(stagedName);
      await waitForEditorLayout(page);

      const pageMeta = await listPageMeta(page);
      const target = pageMeta.find((meta) => meta.pageNumber === pageNumber);
      if (!target) {
        throw new FolioExtractError(
          `folio did not render page ${pageNumber} for ${absoluteDocxPath}`,
        );
      }

      const locator = page.locator(PAGE_SELECTOR).nth(target.domIndex);
      await locator.scrollIntoViewIfNeeded();
      await waitForLayoutStability(page);
      await page.waitForTimeout(STABILITY_SETTLE_MS);

      return await inspectSinglePage(page, target.domIndex);
    } finally {
      await fs.unlink(stagedPath).catch(() => {});
    }
  };

  const close = async (): Promise<void> => {
    await context.close();
    await browser.close();
    if (startedServer && serverProcess) {
      // Kill the process tree we spawned, then sweep the port as a backstop:
      // Vite's grandchild can outlive the tree walk, and killByPort guarantees
      // nothing is left holding our per-worktree port for the next run.
      await killProcessTree(serverProcess.pid);
      await killByPort(PLAYGROUND_PORT);
    }
  };

  return { extract, inspectPage, close };
};
