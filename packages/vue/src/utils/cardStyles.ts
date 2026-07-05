/**
 * Sidebar card chrome. Numeric pixel values rather than `'8px'` strings so both
 * adapters' CSSProperties shapes accept them.
 * @packageDocumentation
 * @public
 */
import type { CSSProperties } from "./cssTypes";

export const CARD_STYLE_COLLAPSED: CSSProperties = {
  padding: "8px 10px",
  borderRadius: 8,
  backgroundColor: "var(--doc-card)",
  cursor: "pointer",
  boxShadow: "var(--doc-card-shadow)",
};

export const CARD_STYLE_EXPANDED: CSSProperties = {
  padding: "10px 12px",
  borderRadius: 8,
  backgroundColor: "var(--doc-surface)",
  cursor: "pointer",
  boxShadow: "var(--doc-card-shadow-strong)",
};
