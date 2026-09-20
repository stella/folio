// Shapes folio-identity-attributes/no-prefix-resolved-identity-read must not flag.
// `scripts/identity-attributes-lint.test.ts` lints this file and asserts zero reports.

declare const element: unknown;
declare const getAttribute: (
  element: unknown,
  namespace: string | null,
  name: string,
) => string | null;
declare const getAttributeAnyPrefix: (element: unknown, name: string) => string | null;
declare const getAttributeByNamespaceUri: (
  element: unknown,
  namespaceUris: ReadonlySet<string>,
  name: string,
) => string | null;
declare const parseNumericAttribute: (
  element: unknown,
  namespace: string | null,
  name: string,
) => number | undefined;
declare const PARA_ID_NAMESPACE_URIS: ReadonlySet<string>;
declare const RELATIONSHIP_NAMESPACE_URIS: ReadonlySet<string>;

// Resolved against the element's namespace scope.
export const paraId = (): string | null =>
  getAttributeByNamespaceUri(element, PARA_ID_NAMESPACE_URIS, "paraId");

export const relationshipId = (): string | null =>
  getAttributeByNamespaceUri(element, RELATIONSHIP_NAMESPACE_URIS, "id");

// Unprefixed: the any-prefix fallback only runs for a prefixed read.
export const packageRelationshipId = (): string | null => getAttribute(element, null, "Id");

// Same local name, different attribute: `w:cols/@w:space` is column spacing,
// and the identity is `xml:space`.
export const columnSpacing = (): number | undefined => parseNumericAttribute(element, "w", "space");

// The complex-script font of `w:rFonts`, not the `r:cs` relationship.
export const complexScriptFont = (): string | null => getAttribute(element, "w", "cs");

// An attribute that describes its own element.
export const value = (): string | null => getAttribute(element, "w", "val");

// The DOM's own method takes a qualified name and resolves nothing.
export const domId = (node: Element): string | null => node.getAttribute("r:id");

// Not a literal: the rule cannot see which attribute this reads.
export const dynamic = (name: string): string | null => getAttributeAnyPrefix(element, name);
