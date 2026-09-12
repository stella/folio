/**
 * Style Engine — explicit cached OOXML style cascade.
 *
 * See {@link createStyleEngine} for the public entry point.
 */

export {
  createStyleEngine,
  type StyleEngine,
  type StyleEngineCacheStats,
  type StyleEngineOptions,
} from "./styleEngine";
export {
  PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS,
  mergeTableParagraphPresentations,
  projectAuthoredParagraphFormatting,
  projectTableParagraphPresentation,
  resolveEffectiveParagraphPresentation,
  type AuthoredParagraphFormatting,
  type EffectiveParagraphPresentation,
  type ParagraphFormattingProjectionDisposition,
  type ParagraphPresentationUnsupportedProperty,
  type ResolvedEffectiveParagraphPresentation,
  type TableParagraphPresentationOverlay,
  type TableParagraphPresentationProjection,
} from "./paragraphPresentation";
export type { ResolvedParagraphStyle } from "../prosemirror/styles/styleResolver";
