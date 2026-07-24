/**
 * Segment-based line fitting seam (premirror port, adapted to folio's
 * measurement idioms).
 *
 * Lineage / credit (moral duty, not just license): the segment-fit design —
 * prepare a text once, then fit lines by arithmetic — is `@chenglou/pretext`'s.
 * The premirror line this seam serves downstream (`@stll/premirror-bridge`,
 * `@premirror/*`) descends from `samwillis/premirror` (MIT © 2026 Sam Willis);
 * those ported packages carry his copyright header and LICENSE. This file is
 * folio-core's own seam — the pluggable extension point — so no concrete engine
 * (and therefore no pretext / premirror code) is imported here.
 *
 * The legacy text walk in `measureParagraph` fits words one canvas
 * measurement at a time and hard-breaks overlong words by binary-searching
 * `measureTextWidth(text.slice(0, mid))` probes (`findMaxFittingLength`) —
 * the probe slices are unique strings, so they pollute the width cache and
 * defeat it (the `workerFontMetrics` prewarm flag exists to soften exactly
 * this). A segment-fit engine replaces that inner loop for plain text runs
 * with a prepare-once model: segments are measured once per (text, font)
 * and lines are then fitted by pure arithmetic (pretext's design).
 *
 * Folio idioms, deliberately mirrored from `measureProvider`:
 * - the engine is a swappable registry (`setSegmentFitEngine` /
 *   `resetSegmentFitEngine`); core never imports a concrete engine —
 *   `@stll/premirror-bridge` installs a `@chenglou/pretext`-backed one;
 * - activation is a feature flag in the `globalThis.__folioFeatureFlags`
 *   bag (`segmentFitLineBreaking`), default OFF — see featureFlags.ts.
 *
 * NOTE (scope): the flag bag and this registry are process-global, like
 * every other folio measurement flag — acceptable while experimental; a
 * per-editor strategy needs options threading before any default flips.
 *
 * Scope guard: this seam replaces ONLY the per-run word walk for plain
 * text runs. Tabs, fields, equations, images, floating zones, justified
 * shrink tolerance, cross-run glue (#991), automatic hyphenation, and CJK
 * dual-font runs stay on the legacy walk (see `styleSupportsSegmentFit` and
 * the call-site admission gate in `measureParagraph`).
 */

import { isSegmentFitLineBreakingEnabled } from "./featureFlags";
import type { FontStyle } from "./measureTypes";

/** One fitted line piece returned by the engine. */
export type SegmentFitLine = {
  /** Exclusive end offset in the run's text (UTF-16 code units). */
  endChar: number;
  /**
   * Measured advance width of the fitted piece in px. MAY include a hung
   * trailing space (pretext does); consumers must not assume either way —
   * the commit path clamps derived trailing-whitespace widths to >= 0.
   */
  width: number;
  /** Opaque continuation cursor; pass back to `fitLine` for the next piece. */
  cursor: unknown;
};

/** A prepared (segmented + measured) text handle. Opaque to core. */
export type SegmentFitPrepared = unknown;

/**
 * A pluggable line-fitting engine consumed by `measureParagraph` when the
 * `segmentFitLineBreaking` flag is on. Implementations prepare a text once
 * (segmenting + measuring) and then fit line pieces by arithmetic.
 */
export type SegmentFitEngine = {
  /**
   * Whether the engine can fit `text` with offsets that stay aligned to the
   * ORIGINAL string. Engines that normalize input before segmenting (e.g.
   * pretext rewrites \r\n, \r, \f) must decline such texts here — otherwise
   * returned end offsets would drift against the run text the painter and
   * click mapping slice. Declined runs measure through the legacy walk.
   */
  supportsText?: (text: string) => boolean;
  /**
   * Prepare `text` for fitting under the given CSS font string (the exact
   * string `buildFontString(style)` produces, so both strategies measure
   * with identical canvas font state). Implementations should cache.
   */
  prepare: (text: string, cssFont: string) => SegmentFitPrepared;
  /**
   * Fit the next line piece starting at `cursor` (null = start of text)
   * into `maxWidth` px. Returns null when nothing fits (caller decides
   * whether to wrap or hand the remainder to the legacy walk). `endChar`
   * must advance strictly beyond the cursor position when a piece is
   * returned.
   */
  fitLine: (
    prepared: SegmentFitPrepared,
    cursor: unknown | null,
    maxWidth: number,
  ) => SegmentFitLine | null;
  /**
   * Drop any prepared/measured state. Invoked via the measuring pipeline's
   * `clearAllCaches()` so engine caches never outlive a font-environment
   * change (web fonts finishing to load) that invalidates the canvas
   * metrics the prepared widths were derived from.
   */
  clearCaches?: () => void;
};

let activeEngine: SegmentFitEngine | null = null;

export const setSegmentFitEngine = (engine: SegmentFitEngine): void => {
  activeEngine = engine;
};

export const resetSegmentFitEngine = (): void => {
  activeEngine = null;
};

export const getSegmentFitEngine = (): SegmentFitEngine | null => activeEngine;

/**
 * Invalidate the installed engine's prepared state. Wired into the
 * measuring pipeline's `clearAllCaches()`.
 */
export const clearSegmentFitEngineCaches = (): void => {
  activeEngine?.clearCaches?.();
};

/** The seam is live only when BOTH the flag is on and an engine is installed. */
export function isSegmentFitActive(): boolean {
  return activeEngine !== null && isSegmentFitLineBreakingEnabled();
}

/**
 * Whether a text run's style can be fitted by the engine without semantic
 * loss. Explicit allowlist over FontStyle members (opt-in gates rot
 * silently, so admit only styles proven engine-safe):
 * - fontFamily, fontSize, bold, italic, fontVariant: carried by
 *   `buildFontString`, engine-safe;
 * - letterSpacing: applied arithmetically outside the font string — decline;
 * - eastAsiaFontFamily: dual-font per-script measurement — decline;
 * - textTransform: legacy transforms text before measuring — decline;
 * - horizontalScale: post-measure width multiplier — decline.
 */
export function styleSupportsSegmentFit(style: FontStyle): boolean {
  return (
    !style.letterSpacing &&
    !style.eastAsiaFontFamily &&
    !style.textTransform &&
    (style.horizontalScale === undefined || style.horizontalScale === 100)
  );
}

/** Callbacks the walk uses to talk to `measureParagraph`'s line state. */
export type SegmentFitWalkHost = {
  /** Current line's remaining budget (includes the caller's tolerance). */
  spaceLeft: () => number;
  lineHasContent: () => boolean;
  /** Commit a fitted piece to the current line (width + exclusive end char). */
  commit: (width: number, endChar: number) => void;
  /** Start a new line at `fromChar` and re-apply the run's font metrics. */
  wrap: (fromChar: number) => void;
};

/**
 * Drive the installed engine over one text run. Returns how many UTF-16
 * code units were consumed; a return < text.length means the engine
 * declined or refused a piece on an empty line (e.g. an overlong token)
 * and the caller's legacy walk must take over from that offset.
 */
export function runSegmentFitWalk(text: string, cssFont: string, host: SegmentFitWalkHost): number {
  const engine = activeEngine;
  if (!engine) {
    return 0;
  }
  if (text.length === 0) {
    return 0;
  }
  if (engine.supportsText && !engine.supportsText(text)) {
    return 0;
  }
  const prepared = engine.prepare(text, cssFont);
  let cursor: unknown | null = null;
  let consumed = 0;
  while (consumed < text.length) {
    // Engines force-fit at least one grapheme per line (CSS behavior);
    // never offer a used-up line, wrap it instead.
    if (host.lineHasContent() && host.spaceLeft() <= 0) {
      host.wrap(consumed);
      continue;
    }
    const piece = engine.fitLine(prepared, cursor, host.spaceLeft());
    if (piece && piece.endChar > consumed) {
      host.commit(piece.width, piece.endChar);
      cursor = piece.cursor;
      consumed = piece.endChar;
      if (consumed < text.length) {
        host.wrap(consumed);
      }
      continue;
    }
    if (host.lineHasContent()) {
      // Nothing fits beside existing content: wrap and retry.
      host.wrap(consumed);
      continue;
    }
    // Engine refused on an empty line (overlong piece): hand the remainder
    // of this run to the caller's legacy hard-breaking.
    break;
  }
  return consumed;
}
