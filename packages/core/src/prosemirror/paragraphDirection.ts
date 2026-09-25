/**
 * Paragraph base-direction modeled as a discriminated union.
 *
 * Replaces the prior `bidi` + `bidiAuto` PM-attribute flag pair, where
 * combinations such as "auto + ltr" or "auto + undecided" were representable
 * but never valid. The states are now mutually exclusive:
 *
 *   - absent (`null`/`undefined`): undecided — default LTR, eligible for
 *     auto-detection.
 *   - `{ source: "auto" }`: auto-detected RTL; re-evaluated as content changes.
 *   - `{ source: "manual"; value }`: an explicit user toggle or imported
 *     `w:bidi`; authoritative, never auto-revisited.
 *
 * The persisted/serialized model keeps the flat OOXML tri-state
 * (`ParagraphFormatting.bidi: boolean | undefined`); `directionToBidi` /
 * `directionFromBidi` bridge the two at the conversion boundary.
 *
 * `source: "auto"` is a *rendering* decision only — it makes the layout
 * painter and the editor's DOM shape the paragraph RTL, so newly typed or
 * pasted directional text displays correctly before the user ever makes an
 * explicit choice. It must never reach the saved package on its own: a
 * paragraph that arrived without `w:bidi` and was never touched by the user
 * must save without `w:bidi` too, no matter how many auto-detection passes
 * ran over it on load. Only `source: "manual"` (an explicit user toggle, or
 * a `w:bidi` the source document already carried) is authored content and
 * belongs in `w:pPr`. Use `directionToBidi` for layout/rendering and
 * `directionToAuthoredBidi` for anything that gets serialized.
 */
export type ParagraphDirection = { source: "auto" } | { source: "manual"; value: "rtl" | "ltr" };

/** Whether the paragraph lays out and exports right-to-left. */
export const directionIsRtl = (direction: ParagraphDirection | null | undefined): boolean =>
  direction?.source === "auto" || (direction?.source === "manual" && direction.value === "rtl");

/**
 * Auto-managed paragraphs (undecided, or previously auto-set) are the ones
 * AutoBidiDetection may (re-)evaluate. A manual decision is left untouched.
 */
export const directionIsAutoManaged = (direction: ParagraphDirection | null | undefined): boolean =>
  direction == null || direction.source === "auto";

/**
 * Map a direction to the OOXML `w:bidi` tri-state for layout/rendering: an
 * auto-detected direction lays out RTL exactly like a manual one. Do not use
 * this to decide what gets written to the saved package — see
 * `directionToAuthoredBidi`.
 */
export const directionToBidi = (
  direction: ParagraphDirection | null | undefined,
): boolean | undefined => {
  if (direction == null) {
    return undefined;
  }
  if (direction.source === "auto") {
    return true;
  }
  return direction.value === "rtl";
};

/**
 * Map a direction to the OOXML `w:bidi` tri-state as *authored* content, for
 * serialization. Only a manual decision (explicit user toggle, or a `w:bidi`
 * the source document already carried) is ever written: `undefined` here
 * means "the paragraph does not state its own direction," which keeps
 * `w:pPr` free of a flag the user never asked for. An auto-detected
 * direction (`source: "auto"`) resolves to `undefined` — same as undecided —
 * regardless of which way detection currently leans, so it stays a view-only
 * aid until the user makes it explicit.
 */
export const directionToAuthoredBidi = (
  direction: ParagraphDirection | null | undefined,
): boolean | undefined => {
  if (direction == null || direction.source === "auto") {
    return undefined;
  }
  return direction.value === "rtl";
};

/**
 * Reconstruct a direction from a model `bidi` value (DOCX import / load): an
 * explicit `true`/`false` is a manual decision, absence is undecided.
 */
export const directionFromBidi = (bidi: boolean | null | undefined): ParagraphDirection | null => {
  if (bidi == null) {
    return null;
  }
  return { source: "manual", value: bidi ? "rtl" : "ltr" };
};

/** Runtime guard for the PM attribute validator. */
export const isParagraphDirection = (value: unknown): value is ParagraphDirection => {
  if (typeof value !== "object" || value === null || !("source" in value)) {
    return false;
  }
  if (value.source === "auto") {
    // `auto` carries no payload; reject a stray `value` so the illegal
    // "auto + ltr/rtl" shape stays unrepresentable.
    return !("value" in value);
  }
  return (
    value.source === "manual" &&
    "value" in value &&
    (value.value === "rtl" || value.value === "ltr")
  );
};
