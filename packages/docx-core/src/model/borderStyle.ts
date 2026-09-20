/**
 * The one place a `CT_Border` `w:val` is decided.
 *
 * `ST_Border` carries two distinct "no border" members. `none` cancels a
 * border the container would otherwise pass down — a table's `w:tblBorders`
 * grid, a paragraph style's `w:pBdr`, a section's `w:pgBorders`; `nil` states
 * that no border is set. Word round-trips whichever the author wrote, and a
 * consumer that folds one into the other loses the cancellation, so both are
 * union members and every consumer asks {@link statesNoBorder} rather than
 * comparing the token. `specifications/reserved-values` records the decision
 * and the lint enforces the ownership.
 *
 * A `w:val` outside the enumeration is kept as {@link UnrecognisedBorderStyle}
 * and written back exactly as it was read, so the union can be closed without
 * the model becoming a place content goes to die.
 */

import { BORDER_STYLES, type BorderStyle } from "./borderStyle.gen";

/**
 * A `w:val` the schema does not declare, kept exactly as the file spelled it.
 *
 * Refusing the border instead would drop an authored edge that Word renders,
 * and collapsing it to a default would rewrite the document on open. The
 * parser reports it through `ParseContext` so the normalisation is visible,
 * and the serializer writes `raw` back unchanged.
 */
export type UnrecognisedBorderStyle = {
  readonly kind: "unrecognised";
  readonly raw: string;
};

/** What `BorderSpec.style` holds: an `ST_Border` member, or the raw token. */
export type BorderStyleValue = BorderStyle | UnrecognisedBorderStyle;

const DECLARED: ReadonlySet<string> = new Set<string>(BORDER_STYLES);

/** Whether the schema declares this token. */
export const isBorderStyle = (value: string): value is BorderStyle => DECLARED.has(value);

/** Read a `w:val` into the model, keeping an undeclared token verbatim. */
export const borderStyleFrom = (value: string): BorderStyleValue =>
  isBorderStyle(value) ? value : { kind: "unrecognised", raw: value };

/** The token to write back, whether or not the schema declares it. */
export const borderStyleToken = (style: BorderStyleValue): string =>
  typeof style === "string" ? style : style.raw;

/** `none`: the author cancelled a border inherited from the container. */
export const isBorderNone = (style: BorderStyleValue | undefined): boolean => style === "none";

/** `nil`: the author stated no border. */
export const isBorderNil = (style: BorderStyleValue | undefined): boolean => style === "nil";

/**
 * Whether the edge paints nothing, by either reserved member.
 *
 * This is what a renderer, a measurer and a collapse rule want: the two
 * members differ in what they do to an inherited border, not in what they
 * draw. Which one was written still matters on the way back out, which is why
 * the value keeps it.
 */
export const statesNoBorder = (style: BorderStyleValue | undefined): boolean =>
  isBorderNone(style) || isBorderNil(style);
