/**
 * Compatibility entry point for the pure tab-measurement helper.
 *
 * Layout owns this calculation. Keep the old ProseMirror subpath as an
 * additive forwarding module for existing consumers.
 */
export {
  calculateSimpleTabWidth,
  calculateTabWidth,
  computeTabStops,
  DEFAULT_TAB_INTERVAL_TWIPS,
  PIXELS_PER_INCH,
  pixelsToTwips,
  TWIPS_PER_INCH,
  twipsToPixels,
  type TabAlignment,
  type TabContext,
  type TabFollowingContent,
  type TabLeader,
  type TabStop,
  type TabWidthResult,
} from "../../layout-engine/measure/tabCalculator";
