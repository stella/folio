/**
 * The parsed XML tree every folio reader shares.
 *
 * The shape is the one `xml-js` produced in non-compact mode, which folio's
 * readers were written against: a root `{ elements }` holding element and
 * text nodes, attributes as a record keyed by their source spelling.
 */
export type XmlElement = {
  declaration?: {
    attributes?: Record<string, string | number>;
  };
  instruction?: string;
  attributes?: Record<string, string | number | undefined>;
  cdata?: string;
  doctype?: string;
  comment?: string;
  text?: string | number | boolean;
  type?: string;
  name?: string;
  /** Resolved namespace URI; non-enumerable so legacy structural comparisons stay stable. */
  namespaceUri?: string;
  /** In-scope namespace declarations; non-enumerable so legacy comparisons stay stable. */
  namespaceScope?: XmlNamespaceScope;
  elements?: XmlElement[];
};

/** Parent-linked XML namespace declarations; each scope stores only local overrides. */
export type XmlNamespaceScope = {
  bindings: ReadonlyMap<string, string>;
  parent?: XmlNamespaceScope;
};
