/**
 * @chenglou/pretext-backed SegmentFitEngine for folio's measuring pipeline
 * (premirror port).
 *
 * prepare() segments + measures text ONCE per (font, text) via pretext
 * (Intl.Segmenter word/grapheme segmentation, per-segment canvas widths,
 * grapheme prefix tables for overflow-wrap). fitLine() is then pure
 * arithmetic — no canvas calls, no string slicing. This replaces the legacy
 * walk's per-call word re-measurement and `findMaxFittingLength`'s
 * slice-probe binary search.
 *
 * Divergences from the legacy walk (both Word/CSS-correct on pretext's
 * side, characterized in pretextParity.test.ts):
 * - trailing whitespace hangs past the line edge instead of forcing a break;
 * - CJK/Thai text breaks between characters (legacy only breaks at
 *   space/hyphen/tab);
 * - overlong tokens char-break from pre-measured grapheme prefix widths.
 *
 * Credit (moral duty, not just license): the segment-fit design and this
 * engine's arithmetic are @chenglou/pretext's (MIT); this file is folio-side
 * glue that adapts pretext to folio-core's `SegmentFitEngine` seam. See
 * NOTICE for the premirror (samwillis/premirror, MIT © Sam Willis) lineage
 * this bridge belongs to.
 */

import { layoutNextLine, prepareWithSegments } from "@chenglou/pretext";
import type {
  SegmentFitEngine,
  SegmentFitLine,
} from "@stll/folio-core/layout-engine/measure/segmentFit";

type LayoutCursor = { segmentIndex: number; graphemeIndex: number };

type PreparedHandle = {
  prepared: ReturnType<typeof prepareWithSegments>;
  segments: string[];
  /** Cumulative UTF-16 offset of each segment start. */
  segCharStart: number[];
  /** Lazily-built grapheme length tables per segment (for mid-segment breaks). */
  graphemeLengths: (number[] | null)[];
};

const PREPARED_CACHE_MAX = 2000;
const preparedCache = new Map<string, PreparedHandle>();

/**
 * U+0000 (NUL) cannot appear in a CSS font string, so it is an unambiguous
 * separator for the composite (font, text) cache key: without it, identical
 * concatenations collide (e.g. font "12px A" + text "B hello" vs font
 * "12px A B" + text "hello").
 */
const KEY_SEP = String.fromCharCode(0);

let graphemeSegmenter: Intl.Segmenter | null = null;
function getGraphemeSegmenter(): Intl.Segmenter {
  if (!graphemeSegmenter) {
    graphemeSegmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
  }
  return graphemeSegmenter;
}

function buildHandle(text: string, cssFont: string): PreparedHandle {
  const prepared = prepareWithSegments(text, cssFont, { whiteSpace: "pre-wrap" });
  const segments = (prepared as unknown as { segments: string[] }).segments;
  const segCharStart: number[] = [];
  let offset = 0;
  for (const segment of segments) {
    segCharStart.push(offset);
    offset += segment.length;
  }
  return {
    prepared,
    segments,
    segCharStart,
    graphemeLengths: Array.from({ length: segments.length }, () => null),
  };
}

function graphemeLengthsFor(handle: PreparedHandle, segmentIndex: number): number[] {
  let lengths = handle.graphemeLengths[segmentIndex];
  if (!lengths) {
    lengths = [];
    const segText = handle.segments[segmentIndex] ?? "";
    for (const g of getGraphemeSegmenter().segment(segText)) {
      lengths.push(g.segment.length);
    }
    handle.graphemeLengths[segmentIndex] = lengths;
  }
  return lengths;
}

function cursorToChar(handle: PreparedHandle, cursor: LayoutCursor): number {
  if (cursor.segmentIndex < 0 || !Number.isFinite(cursor.segmentIndex)) return 0;
  if (cursor.segmentIndex >= handle.segments.length) {
    const last = handle.segments.length - 1;
    if (last < 0) return 0;
    return handle.segCharStart[last]! + handle.segments[last]!.length;
  }
  let chars = handle.segCharStart[cursor.segmentIndex]!;
  if (cursor.graphemeIndex > 0) {
    const lengths = graphemeLengthsFor(handle, cursor.segmentIndex);
    const upto = Math.min(cursor.graphemeIndex, lengths.length);
    for (let i = 0; i < upto; i++) chars += lengths[i]!;
  }
  return chars;
}

/** Exposed for tests: number of prepared handles currently cached. */
export function preparedCacheSize(): number {
  return preparedCache.size;
}

/** Exposed for tests and host teardown. */
export function clearPreparedCache(): void {
  preparedCache.clear();
}

/**
 * Pretext's pre-wrap analysis normalizes CRLF, CR, and FF to LF BEFORE
 * segmenting, so cursor offsets are into the normalized string. Decline any
 * text containing CR or FF so end offsets never drift against the original
 * run (parsed DOCX runs never contain them — breaks are w:br elements — but
 * the seam is public API and must be safe for arbitrary hosts).
 */
const OFFSET_UNSAFE = /[\r\f]/;

export const pretextSegmentFitEngine: SegmentFitEngine = {
  supportsText(text: string): boolean {
    return !OFFSET_UNSAFE.test(text);
  },

  clearCaches(): void {
    clearPreparedCache();
  },

  prepare(text: string, cssFont: string): unknown {
    const key = `${cssFont}${KEY_SEP}${text}`;
    const hit = preparedCache.get(key);
    if (hit) {
      // LRU refresh
      preparedCache.delete(key);
      preparedCache.set(key, hit);
      return hit;
    }
    const handle = buildHandle(text, cssFont);
    preparedCache.set(key, handle);
    if (preparedCache.size > PREPARED_CACHE_MAX) {
      const oldest = preparedCache.keys().next().value;
      if (oldest !== undefined) preparedCache.delete(oldest);
    }
    return handle;
  },

  fitLine(prepared: unknown, cursor: unknown | null, maxWidth: number): SegmentFitLine | null {
    const handle = prepared as PreparedHandle;
    const start: LayoutCursor = (cursor as LayoutCursor | null) ?? {
      segmentIndex: 0,
      graphemeIndex: 0,
    };
    const line = layoutNextLine(handle.prepared, start, maxWidth);
    if (!line) return null;
    const startChar = cursorToChar(handle, line.start);
    const endChar = cursorToChar(handle, line.end);
    if (!Number.isFinite(endChar) || endChar <= startChar) return null;
    return { endChar, width: line.width, cursor: line.end };
  },
};
