/**
 * Shaped text for a PDF writer outside the layout engine: bidirectional
 * levels, the shaper (loaded from its package URL or from the bytes of
 * `@stll/folio-core/text-shaping/wasm`), and the sfnt reader and glyph-id
 * subsetter for embedding the faces shaped glyphs come from.
 *
 * @packageDocumentation
 */

export { parseSfnt, type SfntFont } from "./fonts/sfnt/parse";
export { subsetTrueType, type TrueTypeSubset } from "./fonts/sfnt/subset";
export {
  BIDI_DIRECTION,
  getShaper,
  SHAPING_DIRECTION,
  type BidiLine,
  type ResolveBidiRequest,
  type ShapedRun,
  type Shaper,
  type ShaperSource,
  type ShapeRunRequest,
} from "./shaping/shaper";
