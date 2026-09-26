/**
 * Text shaping for producers outside the layout engine.
 *
 * What a caller needs to set complex-script text into a PDF of its own: the
 * bidirectional algorithm and the shaper (one WebAssembly artifact, loaded from
 * its package URL or from bytes the caller holds), the script segmentation the
 * layout engine selects font slots with, and an sfnt reader and glyph-id
 * subsetter for embedding the faces the shaped glyphs come from.
 *
 * @packageDocumentation
 */

export { parseSfnt, SfntParseError, type SfntFont } from "./fonts/sfnt/parse";
export { SubsetError, subsetTrueType, type TrueTypeSubset } from "./fonts/sfnt/subset";
export {
  BIDI_DIRECTION,
  getShaper,
  ShaperError,
  SHAPING_DIRECTION,
  type BidiDirection,
  type BidiLine,
  type ResolveBidiRequest,
  type ShapedGlyph,
  type ShapedRun,
  type Shaper,
  type ShaperSource,
  type ShapeRunRequest,
  type ShapingDirection,
} from "./shaping/shaper";
export {
  isRightToLeftCodePoint,
  SCRIPT_CLASS,
  segmentByScript,
  type ScriptClass,
  type ScriptSegment,
} from "./utils/scriptSegments";
