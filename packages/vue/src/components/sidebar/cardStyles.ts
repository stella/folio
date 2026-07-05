// Re-export the canonical card chrome from folio's util layer. The util
// `CSSProperties` is the loose framework-neutral shape (string | number
// values); the numeric `borderRadius: 8` entries flow through unchanged.
//
// Upstream re-typed these to Vue's `CSSProperties` via an `as` cast; folio
// forbids `as` casts and no SFC consumes these constants (the cards inline
// their chrome in `<style scoped>`), so the values are re-exported as-is.
export {
  CARD_STYLE_COLLAPSED,
  CARD_STYLE_EXPANDED,
} from "../../utils/cardStyles";
