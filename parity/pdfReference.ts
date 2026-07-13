/**
 * Shared PDF extraction for external DOCX reference renderers.
 *
 * A renderer adapter is responsible only for producing a PDF. This module
 * turns that PDF into the normalized line geometry and page images consumed
 * by the comparison/reporting pipeline.
 */

import { readdir } from "node:fs/promises";
import path from "node:path";

import { CACHE_DIR } from "./config";
import { parseStextXml } from "./stextParse";
import { normalizeLineText } from "./textNorm";
import type { DocGeom, PageGeom } from "./types";

const PNG_DPI = 96;
const PAGE_PNG_RE = /^p(\d+)\.png$/;
const COMPLETE_PAGE_COUNT_FILENAME = ".complete-page-count";

export const sha256OfFile = async (filePath: string): Promise<string> => {
  const data = await Bun.file(filePath).arrayBuffer();
  const hasher = new Bun.CryptoHasher("sha256");
  hasher.update(data);
  return hasher.digest("hex");
};

export const cacheDirFor = (sha256: string): string => path.join(CACHE_DIR, sha256);

/** `mutool -v` prints its version to stderr, not stdout. */
export const getMutoolVersion = async (): Promise<string> => {
  const proc = Bun.spawn(["mutool", "-v"], { stdout: "ignore", stderr: "pipe" });
  const stderr = await new Response(proc.stderr).text();
  await proc.exited;
  return stderr.trim();
};

export const readCachedGeom = async (
  geomPath: string,
  absDocxPath: string,
): Promise<DocGeom | null> => {
  const file = Bun.file(geomPath);
  if (!(await file.exists())) return null;
  const geom = (await file.json()) as DocGeom;
  return Object.assign({}, geom, {
    file: absDocxPath,
    pages: geom.pages.map((page) =>
      Object.assign({}, page, {
        lines: page.lines.map((line) =>
          Object.assign({}, line, { normText: normalizeLineText(line.text) }),
        ),
      }),
    ),
  });
};

export const extractPdfGeometry = async (pdfPath: string, xmlPath: string): Promise<PageGeom[]> => {
  const proc = Bun.spawn(["mutool", "draw", "-F", "stext", "-o", xmlPath, pdfPath], {
    stdout: "ignore",
    stderr: "pipe",
  });
  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  if (exitCode !== 0) {
    throw new Error(
      `mutool stext extraction failed for ${pdfPath} (exit ${exitCode}): ${stderr.trim()}`,
    );
  }

  return parseStextXml(await Bun.file(xmlPath).text());
};

const listPagePngs = async (pagesDir: string): Promise<string[]> => {
  let entries: string[];
  try {
    entries = await readdir(pagesDir);
  } catch {
    return [];
  }
  return entries
    .filter((name) => PAGE_PNG_RE.test(name))
    .sort((a, b) => pageNumberOfPngName(a) - pageNumberOfPngName(b))
    .map((name) => path.join(pagesDir, name));
};

const pageNumberOfPngName = (filename: string): number => {
  const match = PAGE_PNG_RE.exec(filename);
  return match?.[1] === undefined ? 0 : Number(match[1]);
};

const readCompletePageCount = async (pagesDir: string): Promise<number | null> => {
  const marker = Bun.file(path.join(pagesDir, COMPLETE_PAGE_COUNT_FILENAME));
  if (!(await marker.exists())) return null;

  const value = (await marker.text()).trim();
  if (!/^\d+$/u.test(value)) return null;

  const count = Number(value);
  return Number.isSafeInteger(count) ? count : null;
};

type CanReusePagePngCacheOptions = {
  existing: string[];
  completePageCount: number | null;
  maxPages?: number;
};

export const canReusePagePngCache = ({
  existing,
  completePageCount,
  maxPages,
}: CanReusePagePngCacheOptions): boolean => {
  const hasSequentialPages = existing.every(
    (filename, index) => pageNumberOfPngName(path.basename(filename)) === index + 1,
  );
  if (!hasSequentialPages) return false;

  if (completePageCount !== null && existing.length === completePageCount) return true;
  return maxPages !== undefined && existing.length >= maxPages;
};

const renderPagePngs = async (
  pdfPath: string,
  pagesDir: string,
  options: { maxPages?: number },
): Promise<void> => {
  const pageRange = options.maxPages === undefined ? [] : [`1-${options.maxPages}`];
  const proc = Bun.spawn(
    [
      "mutool",
      "draw",
      "-r",
      String(PNG_DPI),
      "-o",
      path.join(pagesDir, "p%d.png"),
      pdfPath,
      ...pageRange,
    ],
    { stdout: "ignore", stderr: "pipe" },
  );
  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  if (exitCode !== 0) {
    throw new Error(
      `mutool PNG rendering failed for ${pdfPath} (exit ${exitCode}): ${stderr.trim()}`,
    );
  }
};

type GetPdfPagePngsOptions = {
  pdfPath: string;
  pagesDir: string;
  maxPages?: number;
};

export const getPdfPagePngs = async ({
  pdfPath,
  pagesDir,
  maxPages,
}: GetPdfPagePngsOptions): Promise<string[]> => {
  const existing = await listPagePngs(pagesDir);
  const completePageCount = await readCompletePageCount(pagesDir);
  if (
    canReusePagePngCache({
      existing,
      completePageCount,
      ...(maxPages === undefined ? {} : { maxPages }),
    })
  ) {
    return maxPages === undefined ? existing : existing.slice(0, maxPages);
  }

  await renderPagePngs(pdfPath, pagesDir, maxPages === undefined ? {} : { maxPages });
  const rendered = await listPagePngs(pagesDir);
  if (maxPages === undefined) {
    await Bun.write(path.join(pagesDir, COMPLETE_PAGE_COUNT_FILENAME), `${rendered.length}\n`);
  }
  return maxPages === undefined ? rendered : rendered.slice(0, maxPages);
};
