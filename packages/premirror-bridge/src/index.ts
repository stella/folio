/**
 * @stll/premirror-bridge — folio-side bridge to the pretext/premirror stack.
 *
 * Exposes the pretext-backed SegmentFitEngine for folio's measurement seam
 * (see `@stll/folio-core` layout-engine/measure/segmentFit.ts). Install at a
 * composition root:
 *
 *   import { pretextSegmentFitEngine } from "@stll/premirror-bridge";
 *   import { setSegmentFitEngine } from "@stll/folio-core/layout-engine/measure/segmentFit";
 *
 *   setSegmentFitEngine(pretextSegmentFitEngine);
 *   globalThis.__folioFeatureFlags = { segmentFitLineBreaking: true };
 *
 * Credit (moral duty): the engine wraps @chenglou/pretext (MIT); this bridge
 * belongs to the premirror line (samwillis/premirror, MIT © Sam Willis). See
 * NOTICE.
 */

export { pretextSegmentFitEngine, clearPreparedCache, preparedCacheSize } from "./pretextEngine";
