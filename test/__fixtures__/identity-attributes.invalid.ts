// Deliberate violations of folio-identity-attributes/no-prefix-resolved-identity-read.
// `scripts/identity-attributes-lint.test.ts` lints this file and asserts the count.

declare const element: unknown;
declare const getAttribute: (
  element: unknown,
  namespace: string | null,
  name: string,
) => string | null;
declare const getAttributeAnyPrefix: (element: unknown, name: string) => string | null;
declare const parseNumericAttribute: (
  element: unknown,
  namespace: string | null,
  name: string,
) => number | undefined;
declare const parseOnOffAttribute: (
  element: unknown,
  namespace: string | null,
  name: string,
) => boolean | undefined;

// A paragraph identity another part joins to, under an extension prefix.
export const paraId = (): string | null => getAttribute(element, "w14", "paraId");

// The same identity read through a second prefix, and through a wrong one:
// `paraId` belongs to no Transitional vocabulary, so every prefix is in scope.
export const threadParent = (): string | null => getAttribute(element, "w15", "paraIdParent");
export const canonicalParaId = (): string | null => getAttribute(element, "w", "paraId");

// A relationship target.
export const embedded = (): string | null => getAttribute(element, "r", "embed");

// An annotation id in the WordprocessingML namespace.
export const annotationId = (): number | undefined => parseNumericAttribute(element, "w", "id");

// A bookmark name, the key a hyperlink anchors to.
export const bookmarkName = (): string | null => getAttribute(element, "w", "name");

// Whitespace significance, where a foreign `:space` would answer instead.
export const preservesSpace = (): boolean | undefined =>
  parseOnOffAttribute(element, "xml", "space");

// The reader that resolves by local name on its own.
export const durableId = (): string | null => getAttributeAnyPrefix(element, "durableId");
