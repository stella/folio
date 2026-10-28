/**
 * The attribute and element names that carry OOXML numeric identities.
 *
 * Kept apart from `numericIds.ts` so the XML reader can retain identity
 * attributes during a scan without importing the validator that reads parts
 * with it.
 */

export const NUMBERING_ATTRIBUTES: ReadonlySet<string> = new Set([
  "numId",
  "abstractNumId",
  "numPicBulletId",
]);
export const ID_VALUE_ELEMENTS: ReadonlySet<string> = new Set([
  "id",
  "numId",
  "abstractNumId",
  "lvlPicBulletId",
]);

export const localName = (name: string): string =>
  name.slice(name.indexOf(":") + 1);

type OoxmlNumericIdAttributeOptions = {
  elementName: string;
  attributeName: string;
};

/** Attributes an identity scan must retain, derived from the domain classifier. */
export const isOoxmlNumericIdAttributeName = ({
  elementName,
  attributeName,
}: OoxmlNumericIdAttributeOptions): boolean => {
  const attribute = localName(attributeName);
  return (
    attribute === "id" ||
    NUMBERING_ATTRIBUTES.has(attribute) ||
    (attribute === "val" && ID_VALUE_ELEMENTS.has(localName(elementName)))
  );
};
