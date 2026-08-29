/**
 * Measurement Cache
 *
 * LRU cache for text width measurements and paragraph layout results.
 * Improves performance by avoiding repeated measurements of identical content.
 */

import type { ImageRun, ParagraphAttrs, ParagraphBlock, ParagraphMeasure, Run } from "../types";
import { getLineBreakProviderGeneration } from "./lineBreakProvider";
import { clearFontResolvedCache } from "./measureHelpers";

// =============================================================================
// TEXT WIDTH CACHE
// =============================================================================

/**
 * Cache entry for text width measurements
 */
type TextWidthEntry = {
  hitsSinceRefresh: number;
  width: number;
};

/**
 * Default max entries for text width cache
 * Large documents (30+ pages) can generate 20,000+ unique text measurements.
 * A generous default avoids cache thrashing on big docs.
 */
const DEFAULT_TEXT_CACHE_SIZE = 20_000;
// Refresh recency in batches. Moving a Map entry on every hit costs two
// mutations; frequently reused measurements remain protected without paying
// that bookkeeping on every line-break probe.
const TEXT_CACHE_REFRESH_AFTER_HITS = 16;

/**
 * Current max size for text width cache
 */
let textCacheMaxSize = DEFAULT_TEXT_CACHE_SIZE;

/**
 * LRU cache for text width measurements
 * Key format: "text|font|letterSpacing"
 */
const textWidthCache = new Map<string, TextWidthEntry>();
let textWidthCacheGeneration = 0;

/**
 * Create a cache key for text width lookup
 */
function makeTextKey(text: string, font: string, letterSpacing: number): string {
  return `${text}|${font}|${letterSpacing || 0}`;
}

/**
 * Evict oldest entries if cache exceeds max size
 */
function evictTextEntries(): void {
  while (textWidthCache.size > textCacheMaxSize) {
    const oldestKey = textWidthCache.keys().next().value;
    if (oldestKey === undefined) {
      break;
    }
    textWidthCache.delete(oldestKey);
  }
}

/**
 * Get cached text width or return undefined
 */
export function getCachedTextWidth(
  text: string,
  font: string,
  letterSpacing: number = 0,
): number | undefined {
  const key = makeTextKey(text, font, letterSpacing);
  const entry = textWidthCache.get(key);

  if (entry !== undefined) {
    entry.hitsSinceRefresh += 1;
    if (entry.hitsSinceRefresh >= TEXT_CACHE_REFRESH_AFTER_HITS) {
      entry.hitsSinceRefresh = 0;
      textWidthCache.delete(key);
      textWidthCache.set(key, entry);
    }
    return entry.width;
  }

  return undefined;
}

/**
 * Store text width in cache
 */
export function setCachedTextWidth(
  text: string,
  font: string,
  letterSpacing: number,
  width: number,
): void {
  const key = makeTextKey(text, font, letterSpacing);
  textWidthCache.set(key, { hitsSinceRefresh: 0, width });
  if (textWidthCache.size > textCacheMaxSize) {
    evictTextEntries();
  }
}

/**
 * Clear the text width cache
 */
export function clearTextWidthCache(): void {
  textWidthCache.clear();
  textWidthCacheGeneration += 1;
}

/**
 * Monotonic generation used by async measurement workers to drop
 * responses that were measured before a cache reset.
 */
export function getTextWidthCacheGeneration(): number {
  return textWidthCacheGeneration;
}

/**
 * Set the maximum size of the text width cache
 */
export function setTextCacheSize(size: number): void {
  if (!Number.isFinite(size) || size <= 0) {
    return;
  }
  textCacheMaxSize = size;
  evictTextEntries();
}

/**
 * Get current text width cache size
 */
export function getTextCacheSize(): number {
  return textWidthCache.size;
}

// =============================================================================
// FONT METRICS CACHE
// =============================================================================

/**
 * Cached font metrics entry
 */
type FontMetricsEntry = {
  ascent: number;
  descent: number;
  lineHeight: number;
  singleLineRatio: number;
};

/**
 * Default max entries for font metrics cache
 */
const DEFAULT_FONT_CACHE_SIZE = 1000;

/**
 * Current max size for font metrics cache
 */
let fontCacheMaxSize = DEFAULT_FONT_CACHE_SIZE;

/**
 * LRU cache for font metrics
 * Key format: "fontFamily|fontSize|bold|italic|fontVariant"
 */
const fontMetricsCache = new Map<string, FontMetricsEntry>();

/**
 * Create a cache key for font metrics lookup
 */
function makeFontKey(
  fontFamily: string,
  fontSize: number,
  bold: boolean = false,
  italic: boolean = false,
  fontVariant?: string,
): string {
  return `${fontFamily}|${fontSize}|${bold}|${italic}|${fontVariant ?? ""}`;
}

/**
 * Evict oldest entries if font cache exceeds max size
 */
function evictFontEntries(): void {
  while (fontMetricsCache.size > fontCacheMaxSize) {
    const oldestKey = fontMetricsCache.keys().next().value;
    if (oldestKey === undefined) {
      break;
    }
    fontMetricsCache.delete(oldestKey);
  }
}

/**
 * Get cached font metrics or return undefined
 */
export function getCachedFontMetrics(
  fontFamily: string,
  fontSize: number,
  bold: boolean = false,
  italic: boolean = false,
  fontVariant?: string,
): FontMetricsEntry | undefined {
  const key = makeFontKey(fontFamily, fontSize, bold, italic, fontVariant);
  const entry = fontMetricsCache.get(key);

  if (entry !== undefined) {
    // Refresh LRU
    fontMetricsCache.delete(key);
    fontMetricsCache.set(key, entry);
    return entry;
  }

  return undefined;
}

/**
 * Store font metrics in cache
 */
export function setCachedFontMetrics(
  fontFamily: string,
  fontSize: number,
  bold: boolean,
  italic: boolean,
  metrics: FontMetricsEntry,
  fontVariant?: string,
): void {
  const key = makeFontKey(fontFamily, fontSize, bold, italic, fontVariant);
  fontMetricsCache.set(key, metrics);
  evictFontEntries();
}

/**
 * Clear the font metrics cache
 */
export function clearFontMetricsCache(): void {
  fontMetricsCache.clear();
}

/**
 * Set the maximum size of the font metrics cache
 */
export function setFontCacheSize(size: number): void {
  if (!Number.isFinite(size) || size <= 0) {
    return;
  }
  fontCacheMaxSize = size;
  evictFontEntries();
}

/**
 * Get current font metrics cache size
 */
export function getFontCacheSize(): number {
  return fontMetricsCache.size;
}

// =============================================================================
// PARAGRAPH MEASURE CACHE
// =============================================================================

/**
 * Cached paragraph measurement entry
 */
type ParagraphMeasureEntry = {
  measure: ParagraphMeasure;
  maxWidth: number;
};

/**
 * Default max entries for paragraph measure cache
 * Large documents can have 500+ unique paragraphs.
 */
const DEFAULT_PARAGRAPH_CACHE_SIZE = 5000;

/**
 * Current max size for paragraph measure cache
 */
let paragraphCacheMaxSize = DEFAULT_PARAGRAPH_CACHE_SIZE;

/**
 * LRU cache for paragraph measurements
 * Key format: block content hash
 */
const paragraphMeasureCache = new Map<string, ParagraphMeasureEntry>();

type ImageMeasureCacheField =
  | "kind"
  | "width"
  | "height"
  | "transform"
  | "wrapType"
  | "displayMode"
  | "distTop"
  | "distBottom"
  | "exactLineHeight"
  | "position";

type ImageIgnoredCacheField =
  | "src"
  | "alt"
  | "opacity"
  | "layoutInCell"
  | "cssFloat"
  | "distLeft"
  | "distRight"
  | "cropTop"
  | "cropRight"
  | "cropBottom"
  | "cropLeft"
  | "borderWidth"
  | "borderColor"
  | "borderStyle"
  | "isInsertion"
  | "isDeletion"
  | "changeAuthor"
  | "changeDate"
  | "changeRevisionId"
  | "pmStart"
  | "pmEnd";

type ClassifiedImageCacheField = ImageMeasureCacheField | ImageIgnoredCacheField;
type ExhaustivelyClassifiedImageRun = [
  Exclude<keyof ImageRun, ClassifiedImageCacheField>,
  Exclude<ClassifiedImageCacheField, keyof ImageRun>,
] extends [never, never]
  ? ImageRun
  : never;

/** Keep image payloads bounded while classifying every field as measured or intentionally ignored. */
const imageMeasureCacheInput = (run: ExhaustivelyClassifiedImageRun) => ({
  kind: run.kind,
  width: run.width,
  height: run.height,
  transform: run.transform,
  wrapType: run.wrapType,
  displayMode: run.displayMode,
  distTop: run.distTop,
  distBottom: run.distBottom,
  exactLineHeight: run.exactLineHeight,
  positioned: run.position !== undefined,
});

const runMeasureCacheInput = (run: Run) => {
  switch (run.kind) {
    case "text":
    case "tab":
    case "lineBreak":
    case "renderedPageBreak":
    case "field":
    case "math":
      return run;
    case "image":
      return imageMeasureCacheInput(run);
    default: {
      const exhaustive: never = run;
      return exhaustive;
    }
  }
};

const paragraphAttrsMeasureCacheInput = (attrs: ParagraphAttrs | undefined) => {
  if (attrs === undefined) return undefined;

  const { snapToGrid, ...measurementAttrs } = attrs;
  if (measurementAttrs.documentGridLinePitch === undefined) {
    return measurementAttrs;
  }
  return { ...measurementAttrs, snapToGrid: snapToGrid !== false };
};

/** Serialize the complete, semantically normalized paragraph measurement contract into a cache key. */
export function hashParagraphBlock(block: ParagraphBlock): string {
  return JSON.stringify({
    lineBreakProviderGeneration: getLineBreakProviderGeneration(),
    attrs: paragraphAttrsMeasureCacheInput(block.attrs),
    runs: block.runs.map(runMeasureCacheInput),
  });
}

/**
 * Evict oldest entries if paragraph cache exceeds max size
 */
function evictParagraphEntries(): void {
  while (paragraphMeasureCache.size > paragraphCacheMaxSize) {
    const oldestKey = paragraphMeasureCache.keys().next().value;
    if (oldestKey === undefined) {
      break;
    }
    paragraphMeasureCache.delete(oldestKey);
  }
}

/**
 * Get cached paragraph measurement or return undefined
 */
export function getCachedParagraphMeasure(
  block: ParagraphBlock,
  maxWidth: number,
): ParagraphMeasure | undefined {
  const key = hashParagraphBlock(block);
  const entry = paragraphMeasureCache.get(key);

  if (entry !== undefined && entry.maxWidth === maxWidth) {
    // Refresh LRU
    paragraphMeasureCache.delete(key);
    paragraphMeasureCache.set(key, entry);
    return entry.measure;
  }

  return undefined;
}

/**
 * Store paragraph measurement in cache
 */
export function setCachedParagraphMeasure(
  block: ParagraphBlock,
  maxWidth: number,
  measure: ParagraphMeasure,
): void {
  const key = hashParagraphBlock(block);
  paragraphMeasureCache.set(key, { measure, maxWidth });
  evictParagraphEntries();
}

/**
 * Clear the paragraph measure cache
 */
export function clearParagraphMeasureCache(): void {
  paragraphMeasureCache.clear();
}

/**
 * Set the maximum size of the paragraph measure cache
 */
export function setParagraphCacheSize(size: number): void {
  if (!Number.isFinite(size) || size <= 0) {
    return;
  }
  paragraphCacheMaxSize = size;
  evictParagraphEntries();
}

/**
 * Get current paragraph measure cache size
 */
export function getParagraphCacheSize(): number {
  return paragraphMeasureCache.size;
}

// =============================================================================
// GLOBAL CACHE MANAGEMENT
// =============================================================================

/**
 * Clear all measurement caches
 * Call when fonts change, page width changes, or for testing
 */
export function clearAllCaches(): void {
  clearTextWidthCache();
  clearFontMetricsCache();
  clearParagraphMeasureCache();
  clearFontResolvedCache();
}

/**
 * Get total size of all caches
 */
export function getTotalCacheSize(): number {
  return getTextCacheSize() + getFontCacheSize() + getParagraphCacheSize();
}
