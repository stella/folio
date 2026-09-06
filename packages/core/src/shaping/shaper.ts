/**
 * The boundary to the text shaper.
 *
 * Shaping is the one part of text layout that cannot be derived from code
 * points: Arabic letters take their form from their neighbours, lam-alef
 * ligates, Devanagari reorders and forms conjuncts, Hebrew points hang off the
 * letter they belong to. One implementation answers that question, in Rust, and
 * every caller reaches it through this module: a measurement in CSS pixels and
 * a PDF text matrix scale the same glyph ids and the same advances, so the two
 * cannot describe different text.
 *
 * The artifact is loaded on demand. It carries the OpenType layout engine and
 * the Unicode data that engine needs, and a document in Latin, Cyrillic or
 * Greek needs none of it: `getShaper` is called only once a run in a shaping
 * script is about to be measured or painted, and never before.
 */

import { TaggedError } from "better-result";

import initializeShaper, { shapeRun as shapeRunInWasm } from "../generated/text_shaper.js";

/** Which way a run advances. Resolved by the producer, never guessed here. */
export const SHAPING_DIRECTION = {
  leftToRight: "ltr",
  rightToLeft: "rtl",
} as const;

export type ShapingDirection = (typeof SHAPING_DIRECTION)[keyof typeof SHAPING_DIRECTION];

export type ShapeRunRequest = {
  /** sfnt bytes of the face to shape with, which must be the face measured. */
  readonly font: Uint8Array;
  /** Index within a font collection; 0 for a plain font file. */
  readonly faceIndex?: number;
  readonly text: string;
  readonly direction: ShapingDirection;
  /** ISO-15924 tag, or omitted to let the shaper detect the script. */
  readonly script?: string;
  /** BCP-47 language, or omitted. Selects locale-specific substitutions. */
  readonly language?: string;
  /** OpenType features to force on, such as `liga` or `kern`. */
  readonly featuresOn?: readonly string[];
  /** Features to force off. */
  readonly featuresOff?: readonly string[];
};

/** One shaped glyph, in the face's design units. */
export type ShapedGlyph = {
  /** Glyph id in the face that was shaped, not a code point. */
  readonly glyphId: number;
  /**
   * UTF-8 byte offset into the request's text of the character this glyph came
   * from. A ligature's glyphs share one offset, which is what lets a caller map
   * a glyph back to text for extraction and for a `/ToUnicode` map.
   */
  readonly cluster: number;
  readonly xAdvance: number;
  readonly yAdvance: number;
  readonly xOffset: number;
  readonly yOffset: number;
};

export type ShapedRun = {
  readonly glyphs: readonly ShapedGlyph[];
  /** Design units per em, so a caller scales without re-reading the face. */
  readonly unitsPerEm: number;
};

export type Shaper = {
  readonly shapeRun: (request: ShapeRunRequest) => ShapedRun;
};

export class ShaperError extends TaggedError("ShaperError")<{
  message: string;
  cause?: unknown;
}> {}

/**
 * The flat buffer the WebAssembly entry returns:
 * `[unitsPerEm, glyphCount, (glyphId, cluster, xAdvance, yAdvance, xOffset,
 * yOffset) * glyphCount]`. A page of Arabic is tens of thousands of glyphs, and
 * an object each would cost more in boundary crossings than the shaping itself.
 */
const HEADER_FIELDS = 2;
const FIELDS_PER_GLYPH = 6;

const decode = (buffer: Int32Array): ShapedRun => {
  const unitsPerEm = buffer[0] ?? 0;
  const glyphCount = buffer[1] ?? 0;
  const glyphs: ShapedGlyph[] = [];
  for (let index = 0; index < glyphCount; index += 1) {
    const at = HEADER_FIELDS + index * FIELDS_PER_GLYPH;
    glyphs.push({
      glyphId: buffer[at] ?? 0,
      cluster: buffer[at + 1] ?? 0,
      xAdvance: buffer[at + 2] ?? 0,
      yAdvance: buffer[at + 3] ?? 0,
      xOffset: buffer[at + 4] ?? 0,
      yOffset: buffer[at + 5] ?? 0,
    });
  }
  return { glyphs, unitsPerEm };
};

const shaper: Shaper = {
  shapeRun: ({
    font,
    faceIndex = 0,
    text,
    direction,
    script = "",
    language = "",
    featuresOn = [],
    featuresOff = [],
  }) => {
    try {
      return decode(
        shapeRunInWasm(
          font,
          faceIndex,
          text,
          direction === SHAPING_DIRECTION.rightToLeft,
          script,
          language,
          [...featuresOn],
          [...featuresOff],
        ),
      );
    } catch (cause) {
      throw new ShaperError({ message: "Could not shape a run of text", cause });
    }
  },
};

let loading: Promise<Shaper> | undefined;

/**
 * Loads the shaper, once, and hands back the single implementation both the
 * measurer and the PDF backend use.
 *
 * Calling this is what fetches the artifact, so a caller asks for it only after
 * establishing that some run actually needs shaping.
 */
export const getShaper = (): Promise<Shaper> => {
  loading ??= initializeShaper()
    .then(() => {
      resolved = shaper;
      return shaper;
    })
    .catch((cause: unknown) => {
      // A failed load must not poison every later attempt: a transient fetch
      // failure is not a permanent absence of a shaper.
      loading = undefined;
      throw new ShaperError({ message: "Could not initialize the text shaper", cause });
    });
  return loading;
};

/**
 * Whether the artifact has been loaded in this process.
 *
 * Exists for the test that a document in a non-shaping script never pays for
 * it, which is the whole reason this is a separate artifact.
 */
export const shaperLoaded = (): boolean => loading !== undefined;

let resolved: Shaper | null = null;

/**
 * The shaper if it is already loaded, and never a load.
 *
 * The measure seam is synchronous: a measurement cannot await an artifact. So
 * whoever is about to measure text that shapes resolves {@link getShaper}
 * first, and the measurer reads the result here.
 */
export const resolvedShaper = (): Shaper | null => resolved;
