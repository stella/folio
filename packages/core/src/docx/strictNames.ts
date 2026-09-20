/**
 * Reading a slot whose name a Strict producer spells by writing direction.
 *
 * ECMA-376 Part 1 names a horizontal edge `start`/`end` where Part 4 names it
 * `left`/`right`, and Part 4 declares both so a Transitional consumer reads
 * either. folio's model holds one direction and every save writes the physical
 * name, so the tolerance belongs on the way in — once, over the generated
 * table, rather than as a `?? findChild(…, "start")` at each reader. A slot the
 * table does not name is not a renamed slot, and passing one here is a compile
 * error rather than a lookup that silently finds nothing.
 */

import { type RenamedSlot, STRICT_NAMES_BY_TRANSITIONAL_NAME } from "./strictNames.gen";
import { findChild, getAttribute, parseNumericAttribute, type XmlElement } from "./xmlParser";

/**
 * The spellings a document may have used for a slot, folio's own first.
 *
 * The key carries the complex type, so the local name is what follows the
 * space, without the `@` that marks an attribute.
 */
const spellingsOf = (slot: RenamedSlot): readonly string[] => {
  const local = slot.slice(slot.indexOf(" ") + 1);
  const transitional = local.startsWith("@") ? local.slice(1) : local;
  return [transitional, ...STRICT_NAMES_BY_TRANSITIONAL_NAME[slot]];
};

/** A renamed child of `parent`, in whichever spelling the document used. */
export const findChildAnySpelling = (
  parent: XmlElement | null | undefined,
  slot: RenamedSlot,
): XmlElement | null => {
  for (const spelling of spellingsOf(slot)) {
    const child = findChild(parent, "w", spelling);
    if (child !== null) {
      return child;
    }
  }
  return null;
};

/** A renamed attribute's number, in whichever spelling the document used. */
export const numericAttributeAnySpelling = (
  element: XmlElement | null | undefined,
  slot: RenamedSlot,
): number | undefined => {
  for (const spelling of spellingsOf(slot)) {
    const value = parseNumericAttribute(element, "w", spelling);
    if (value !== undefined) {
      return value;
    }
  }
  return undefined;
};

/** Whether the element carries a renamed attribute under either spelling. */
export const hasAttributeAnySpelling = (
  element: XmlElement | null | undefined,
  slot: RenamedSlot,
): boolean => spellingsOf(slot).some((spelling) => getAttribute(element, "w", spelling) !== null);
