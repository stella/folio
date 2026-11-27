/**
 * Where automatic hyphenation gets its dictionaries. They ship as data files
 * (`@stll/folio-core/hyphenation/<id>.json`) fetched the first time a
 * document hyphenates in their language; an editor lays the document out
 * again when one arrives, and without one it does not hyphenate.
 *
 * A host that cannot fetch them from beside the package points
 * `configureHyphenationAssets` at another location, or hands over the parsed
 * files. `preloadHyphenationDictionaries` loads them ahead of a layout that
 * must hyphenate on its first pass.
 *
 * @packageDocumentation
 */

export {
  configureHyphenationAssets,
  HYPHENATION_DICTIONARY_IDS,
  HyphenationDictionaryError,
  hyphenationDictionaryFor,
  type HyphenationAssetOptions,
  type HyphenationDictionaryId,
  type HyphenationPatterns,
} from "./layout-engine/measure/hyphenationDictionaries";
export { preloadHyphenationDictionaries } from "./layout-engine/measure/hyphenationPreload";
