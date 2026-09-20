/**
 * The one place an `a:ln/a:prstDash` `@val` is decided.
 *
 * `ST_PresetLineDashVal` is DrawingML's own dash vocabulary. Three of its
 * spellings (`solid`, `dash`, `dot`) also occur in CSS `border-style` and in
 * `text-decoration-style`, where they name different patterns, so an outline
 * dash is kept in its own type and resolved against its own table rather than
 * matched by string against whatever vocabulary a consumer happens to hold.
 *
 * A `@val` outside the enumeration is kept as {@link UnrecognisedPresetLineDash}
 * and written back exactly as it was read, the way an undeclared `ST_Border`
 * `w:val` is.
 *
 * `a:custDash` is not this attribute and is not modelled: an outline that
 * carries one keeps it through `ShapeOutline.rawXml`, which the serializer
 * replays verbatim, so the custom stop list survives a save without the dash
 * field pretending to describe it.
 */

import { PRESET_LINE_DASH_VALS, type PresetLineDashVal } from "./presetLineDash.gen";

/**
 * A `@val` the schema does not declare, kept exactly as the file spelled it.
 *
 * Dropping the dash would repaint an authored outline as a plain line, and
 * reading it as `solid` would rewrite the document on open. The parser reports
 * it through `ParseContext` so the normalisation is visible, and the
 * serializer writes `raw` back unchanged.
 */
export type UnrecognisedPresetLineDash = {
  readonly kind: "unrecognised";
  readonly raw: string;
};

/** What `ShapeOutline.dash` holds: a schema member, or the raw token. */
export type PresetLineDashValue = PresetLineDashVal | UnrecognisedPresetLineDash;

const DECLARED: ReadonlySet<string> = new Set<string>(PRESET_LINE_DASH_VALS);

/** Whether the schema declares this token. */
export const isPresetLineDashVal = (value: string): value is PresetLineDashVal =>
  DECLARED.has(value);

/** Read a `@val` into the model, keeping an undeclared token verbatim. */
export const presetLineDashFrom = (value: string): PresetLineDashValue =>
  isPresetLineDashVal(value) ? value : { kind: "unrecognised", raw: value };

/** The token to write back, whether or not the schema declares it. */
export const presetLineDashToken = (dash: PresetLineDashValue): string =>
  typeof dash === "string" ? dash : dash.raw;
