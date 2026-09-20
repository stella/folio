// Fixture for `folio-xml-splice/no-hand-rolled-splice`. Applying the
// replacements through the owner is the accepted spelling; a concatenation
// that deletes nothing is out of the rule's scope.

// `spliceXml` is module-local to `packages/core/src/docx/selectiveXmlPatch.ts`,
// so the fixture names its shape rather than importing it.
declare const spliceXml: (
  xml: string,
  splices: readonly { start: number; end: number; newXml: string }[],
) => string | null;

export const throughTheOwner = (
  xml: string,
  start: number,
  end: number,
  newXml: string,
): string | null => spliceXml(xml, [{ start, end, newXml }]);

/** One slice cannot cut a region out: this appends before the root close. */
export const append = (xml: string, rootClose: number, definitions: string): string =>
  xml.slice(0, rootClose) + definitions;

/** Assembled from parts already extracted, so the chain slices nothing. */
export const fromParts = (head: string, minted: string, middle: string, tail: string): string =>
  head + minted + middle + tail;

/** Two slices of different strings are two reads, not one splice. */
export const joinTwoParts = (first: string, second: string, at: number): string =>
  first.slice(0, at) + second.slice(at);
