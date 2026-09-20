/**
 * `@w:fldLock` and `@w:dirty`, the two `ST_OnOff` attributes every field form
 * carries.
 *
 * A simple field states them on `<w:fldSimple>`; a complex field states them on
 * the `<w:fldChar w:fldCharType="begin"/>` that opens it, and folio hoists them
 * from there onto the assembled `ComplexField`. Neither attribute has an XSD
 * default, so a field states one of three things about each: nothing, an
 * explicit off, an explicit on. `w:dirty="0"` on a field inside a `TOC` result
 * is the case that matters: it says "do not recompute this", which is not what
 * an absent attribute says.
 *
 * Three readers and three writers had drifted into the same `=== true` /
 * truthiness collapse, so both directions live here instead.
 */

import { parseOnOffAttribute, type XmlElement } from "./xmlParser";

/** What a field element states about its lock and dirty flags. */
export type FieldState = {
  fldLock?: boolean;
  dirty?: boolean;
};

/** Read both attributes off a `w:fldSimple` or `w:fldChar`, keeping an explicit off. */
export const parseFieldState = (element: XmlElement): FieldState => {
  const state: FieldState = {};
  const fldLock = parseOnOffAttribute(element, "w", "fldLock");
  if (fldLock !== undefined) {
    state.fldLock = fldLock;
  }
  const dirty = parseOnOffAttribute(element, "w", "dirty");
  if (dirty !== undefined) {
    state.dirty = dirty;
  }
  return state;
};

/**
 * The stated flags of a field, without the rest of it.
 *
 * The complex-field assembly carries the begin `w:fldChar`'s state across the
 * runs between `begin` and `end`, and this is the one place that knows which
 * fields that state is made of.
 */
export const fieldStateOf = ({ fldLock, dirty }: FieldState): FieldState => {
  const state: FieldState = {};
  if (fldLock !== undefined) {
    state.fldLock = fldLock;
  }
  if (dirty !== undefined) {
    state.dirty = dirty;
  }
  return state;
};

/**
 * The attributes to write for a field's stated flags, in schema order.
 *
 * `1`/`0` is what Word writes and what the table, section and border
 * serializers write.
 */
export const fieldStateAttributes = ({ fldLock, dirty }: FieldState): string[] => {
  const attrs: string[] = [];
  if (fldLock !== undefined) {
    attrs.push(`w:fldLock="${fldLock ? "1" : "0"}"`);
  }
  if (dirty !== undefined) {
    attrs.push(`w:dirty="${dirty ? "1" : "0"}"`);
  }
  return attrs;
};
