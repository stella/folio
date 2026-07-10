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

export type FolioExtractor = {
  extract: (docxPath: string, options?: FolioExtractOptions) => Promise<FolioExtraction>;
  inspectPage: (docxPath: string, pageNumber: number) => Promise<FolioPageInspection>;
  close: () => Promise<void>;
};

const EDITOR_SELECTOR = '[data-testid="folio-editor"]';
const PAGE_SELECTOR = ".layout-page";

const VIEWPORT = { width: 1400, height: 1000 };

const SERVER_PROBE_TIMEOUT_MS = 2000;
const SERVER_START_TIMEOUT_MS = 90_000;
const SERVER_POLL_INTERVAL_MS = 500;

const PLAYGROUND_NAVIGATION_TIMEOUT_MS = 120_000;
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
  /** Stable page-local table-cell identity, when the line is inside a cell. */
  visualGroup?: string;
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
      ...(rawLine.visualGroup !== undefined ? { visualGroup: rawLine.visualGroup } : {}),
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

const installCleanScreenshotStyle = async (page: Page): Promise<void> => {
  await page.addStyleTag({
    content: `
      .pg-collab-header,
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
    `,
  });
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
  page.evaluate<RawPage, { domIndex: number; meaningfulInkPattern: string }>(
    (input) => {
      const { domIndex: idx, meaningfulInkPattern } = input;
      const meaningfulInkCharacter = new RegExp(meaningfulInkPattern, "u");
      const pageEls = Array.from(document.querySelectorAll(".layout-page")) as HTMLElement[];
      const el = pageEls[idx];
      if (!el) {
        throw new Error(`layout-page at domIndex ${idx} disappeared`);
      }

      const pageRect = el.getBoundingClientRect();
      const dataPageNumber = Number(el.dataset["pageNumber"]);
      const pageNumber = Number.isFinite(dataPageNumber) ? dataPageNumber : idx + 1;

      const lineEls = Array.from(el.querySelectorAll(".layout-line")) as HTMLElement[];
      const lines = lineEls.flatMap((lineEl) => {
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
        const region = (() => {
          if (lineEl.closest(".layout-page-header")) {
            return "header";
          }
          if (lineEl.closest(".layout-page-footer")) {
            return "footer";
          }
          return "body";
        })();
        const tableCell = lineEl.closest(".layout-table-cell");
        const visualGroup = tableCell
          ? `table-cell:${Array.from(el.querySelectorAll(".layout-table-cell")).indexOf(tableCell)}`
          : undefined;

        const fontFrom = (sourceEl: HTMLElement | null) => {
          if (!sourceEl) return {};
          const computed = getComputedStyle(sourceEl);
          const parsedSize = Number.parseFloat(computed.fontSize);
          return {
            fontFamilyRaw: computed.fontFamily,
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

        const toRawLine = (text: string, rect: DOMRect, sourceEl: HTMLElement | null) => ({
          text,
          rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
          region,
          ...(visualGroup !== undefined ? { visualGroup } : {}),
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

        const runEl = lineEl.querySelector(".layout-run");

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
    { domIndex, meaningfulInkPattern: MEANINGFUL_INK_CHARACTER.source },
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
      const region = (() => {
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
  opts: { headless?: boolean; reuseServer?: boolean } = {},
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
  const serverAlreadyUp =
    reuseServer && (await probeServer(PLAYGROUND_URL, SERVER_PROBE_TIMEOUT_MS));
  if (!serverAlreadyUp) {
    // Clear any leftover server on OUR per-worktree port (see killByPort:
    // port-scoped, so other worktrees' servers are untouched), then start
    // fresh, telling the playground which port to bind via env.
    await killByPort(PLAYGROUND_PORT);
    serverProcess = Bun.spawn(PLAYGROUND_DEV_COMMAND, {
      cwd: REPO_ROOT,
      env: { ...process.env, FOLIO_PLAYGROUND_PORT: String(PLAYGROUND_PORT) },
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
      await killByPort(PLAYGROUND_PORT);
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
      await page.goto(`${PLAYGROUND_URL}/?file=${encodeURIComponent(stagedName)}`, {
        waitUntil: "domcontentloaded",
        timeout: PLAYGROUND_NAVIGATION_TIMEOUT_MS,
      });
      await waitForEditorLayout(page);

      const pageMeta = await listPageMeta(page);
      if (pageMeta.length === 0) {
        throw new FolioExtractError(`folio rendered zero pages for ${absoluteDocxPath}`);
      }
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
      await page.goto(`${PLAYGROUND_URL}/?file=${encodeURIComponent(stagedName)}`, {
        waitUntil: "domcontentloaded",
        timeout: PLAYGROUND_NAVIGATION_TIMEOUT_MS,
      });
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
