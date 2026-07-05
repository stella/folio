/**
 * Style Engine
 *
 * Explicit, cached cascade resolution over OOXML style definitions.
 * Wraps the existing {@link StyleResolver} (which encodes the cascade
 * order per ECMA-376 §17.7) with per-method memoization so that callers
 * walking many paragraphs that reference the same handful of styleIds
 * pay the resolution cost once per distinct key.
 *
 * Cascade order (lowest → highest precedence, per ECMA-376 §17.7.2):
 *   docDefaults → default-of-type style → linked character style →
 *   named style chain (basedOn) → direct formatting.
 *
 * The engine does not change cascade semantics; it is a pure cache in
 * front of the legacy resolver, kept additive so callers can be migrated
 * one at a time.
 */

import type {
  ResolvedParagraphStyle,
  StyleResolver,
  TableCellParagraphSpacingOverlay,
} from "../prosemirror/styles/styleResolver";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import type { DocDefaults, Style, StyleDefinitions, TextFormatting } from "../types/document";

/**
 * Sentinel key used for null / undefined styleId queries so that
 * {@link Map} lookups distinguish "cached an undefined result" from
 * "never queried" without storing JS `undefined` as a key.
 */
const DEFAULT_KEY = "\0__default__";

function cacheKey(styleId: string | undefined | null): string {
  if (styleId === undefined || styleId === null || styleId === "") {
    return DEFAULT_KEY;
  }
  return styleId;
}

/** Engine construction options. */
export type StyleEngineOptions = {
  /**
   * Toggle the memoization layer.
   *
   * Defaults to `true`. Set to `false` to force every call through the
   * underlying {@link StyleResolver}; useful in tests where stale-cache
   * bugs could mask real regressions, and for diagnosing cache-related
   * behaviour without rebuilding the engine.
   */
  cache?: boolean;
};

/** Lightweight observability for cache effectiveness. */
export type StyleEngineCacheStats = {
  hits: number;
  misses: number;
  /** Total entries across all internal caches. */
  size: number;
};

/**
 * Cached, explicit style cascade engine.
 *
 * The engine exposes the same surface as {@link StyleResolver} for the
 * methods folio currently consumes, plus cache controls. Returned
 * objects are shared between cache hits — treat them as immutable.
 */
export type StyleEngine = {
  /** Look up a single named style by id. */
  getStyle: (styleId: string) => Style | undefined;
  /** True if a style with the given id is registered. */
  hasStyle: (styleId: string) => boolean;
  /** Return the document-wide defaults (`w:docDefaults`). */
  getDocDefaults: () => DocDefaults | undefined;
  /** The style flagged `w:default="1"` for paragraphs, else "Normal". */
  getDefaultParagraphStyle: () => Style | undefined;
  /** The style flagged `w:default="1"` for character styles. */
  getDefaultCharacterStyle: () => Style | undefined;
  /** The style flagged `w:default="1"` for tables. */
  getDefaultTableStyle: () => Style | undefined;
  /** All visible paragraph styles, sorted for toolbar use. */
  getParagraphStyles: () => Style[];
  /** All visible table styles, sorted for the table-style gallery. */
  getTableStyles: () => Style[];

  /**
   * Resolve the full paragraph cascade for a given styleId.
   * Returned object is cached and must not be mutated.
   */
  resolveParagraphStyle: (styleId: string | undefined | null) => ResolvedParagraphStyle;
  /**
   * Resolve the paragraph cascade for a paragraph inside a table cell,
   * layering the enclosing table style's paragraph-spacing fields between
   * docDefaults and the paragraph's own style chain — see
   * {@link StyleResolver.resolveParagraphStyleInTable}.
   *
   * Pass the *same* `tableParagraphOverlay` object reference for every
   * paragraph governed by the same table style region (folio's table
   * conversion already computes it once per region) — the cache below keys
   * on object identity, so a fresh object per call defeats memoization.
   * Returned object is cached and must not be mutated.
   */
  resolveParagraphStyleInTable: (
    styleId: string | undefined | null,
    tableParagraphOverlay: TableCellParagraphSpacingOverlay | undefined,
  ) => ResolvedParagraphStyle;
  /**
   * Resolve a run/character style with docDefaults applied.
   * Returns `undefined` when the cascade yields nothing.
   */
  resolveRunStyle: (styleId: string | undefined | null) => TextFormatting | undefined;
  /**
   * Return a run style's own properties without docDefaults applied —
   * used when the caller has already merged docDefaults via the
   * paragraph cascade and would otherwise double-apply font defaults.
   */
  getRunStyleOwnProperties: (styleId: string | undefined | null) => TextFormatting | undefined;

  /** Drop every memoized cascade result. */
  invalidate: () => void;
  /** Inspect cache effectiveness (hits, misses, current size). */
  stats: () => StyleEngineCacheStats;
};

/**
 * Build a {@link StyleEngine} from parsed style definitions.
 *
 * @param styleDefinitions - The parsed `styles.xml` package, or `undefined`
 *   for documents without a styles part.
 * @param options - Engine-level toggles, see {@link StyleEngineOptions}.
 */
export function createStyleEngine(
  styleDefinitions: StyleDefinitions | undefined,
  options?: StyleEngineOptions,
): StyleEngine {
  const resolver: StyleResolver = createStyleResolver(styleDefinitions);
  const cacheEnabled = options?.cache ?? true;

  // Wrap values in a single-property record so that `Map.get` returning
  // `undefined` unambiguously means "not cached" — even when the cached
  // value itself is `undefined` (e.g., resolveRunStyle for a styleId
  // whose cascade resolves to nothing). Avoids the unsafe `as T` cast
  // we would otherwise need on the `has` / `get` boundary.
  type Cached<T> = { value: T };
  const paragraphCache = new Map<string, Cached<ResolvedParagraphStyle>>();
  const runCache = new Map<string, Cached<TextFormatting | undefined>>();
  const ownPropsCache = new Map<string, Cached<TextFormatting | undefined>>();
  // Keyed by the overlay object's identity (not its contents) since table
  // conversion computes one overlay object per style region and reuses it
  // across every cell/paragraph in that region — see the interface doc
  // comment on `resolveParagraphStyleInTable`. A `let` (not `const`) so
  // `invalidate()` can drop every entry without a `WeakMap.clear()`, which
  // doesn't exist.
  let tableOverlayCache = new WeakMap<
    TableCellParagraphSpacingOverlay,
    Map<string, Cached<ResolvedParagraphStyle>>
  >();

  let hits = 0;
  let misses = 0;

  const memoize = <T>(cache: Map<string, Cached<T>>, key: string, compute: () => T): T => {
    if (!cacheEnabled) {
      misses += 1;
      return compute();
    }
    const cached = cache.get(key);
    if (cached !== undefined) {
      hits += 1;
      return cached.value;
    }
    misses += 1;
    const value = compute();
    cache.set(key, { value });
    return value;
  };

  const memoizeInTable = (
    styleId: string | undefined | null,
    tableParagraphOverlay: TableCellParagraphSpacingOverlay,
  ): ResolvedParagraphStyle => {
    if (!cacheEnabled) {
      misses += 1;
      return resolver.resolveParagraphStyleInTable(styleId, tableParagraphOverlay);
    }
    let overlayCache = tableOverlayCache.get(tableParagraphOverlay);
    if (!overlayCache) {
      overlayCache = new Map();
      tableOverlayCache.set(tableParagraphOverlay, overlayCache);
    }
    return memoize(overlayCache, cacheKey(styleId), () =>
      resolver.resolveParagraphStyleInTable(styleId, tableParagraphOverlay),
    );
  };

  return {
    getStyle: (styleId) => resolver.getStyle(styleId),
    hasStyle: (styleId) => resolver.hasStyle(styleId),
    getDocDefaults: () => resolver.getDocDefaults(),
    getDefaultParagraphStyle: () => resolver.getDefaultParagraphStyle(),
    getDefaultCharacterStyle: () => resolver.getDefaultCharacterStyle(),
    getDefaultTableStyle: () => resolver.getDefaultTableStyle(),
    getParagraphStyles: () => resolver.getParagraphStyles(),
    getTableStyles: () => resolver.getTableStyles(),

    resolveParagraphStyle: (styleId) =>
      memoize(paragraphCache, cacheKey(styleId), () => resolver.resolveParagraphStyle(styleId)),
    resolveParagraphStyleInTable: (styleId, tableParagraphOverlay) =>
      tableParagraphOverlay === undefined
        ? memoize(paragraphCache, cacheKey(styleId), () =>
            resolver.resolveParagraphStyleInTable(styleId, undefined),
          )
        : memoizeInTable(styleId, tableParagraphOverlay),
    resolveRunStyle: (styleId) =>
      memoize(runCache, cacheKey(styleId), () => resolver.resolveRunStyle(styleId)),
    getRunStyleOwnProperties: (styleId) =>
      memoize(ownPropsCache, cacheKey(styleId), () => resolver.getRunStyleOwnProperties(styleId)),

    invalidate: () => {
      paragraphCache.clear();
      runCache.clear();
      ownPropsCache.clear();
      // WeakMap has no `.clear()` — drop the whole map so stale per-overlay
      // caches from a previous style resolution can't leak forward.
      tableOverlayCache = new WeakMap();
      hits = 0;
      misses = 0;
    },
    stats: () => ({
      hits,
      misses,
      // Table-overlay entries live in a WeakMap keyed by overlay object
      // identity and aren't enumerable, so they're intentionally excluded
      // from this count — it undercounts total cache size whenever
      // resolveParagraphStyleInTable has been used with an overlay.
      size: paragraphCache.size + runCache.size + ownPropsCache.size,
    }),
  };
}
