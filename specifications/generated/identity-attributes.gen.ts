/**
 * GENERATED FILE - do not edit.
 *
 * The OOXML attributes whose value names something outside the element that
 * carries it, derived from the Transitional schema graph and Word's extension
 * vocabularies by `scripts/generate-identity-attributes.ts`. Regenerate with:
 *
 *   bun run generate:identity-attributes
 */

/** An attribute a prefix-resolved read can join to the wrong thing. */
export type IdentityAttribute = {
  /** The prefixes the identity is read through, or `null` for every prefix. */
  readonly prefixes: readonly string[] | null;
  readonly reason: string;
  readonly source: "extension" | "schema";
};

/**
 * One attribute per line: local name, a tab, the prefixes it is read through
 * (comma-separated, or `*` for every prefix), a tab, where it came from, a
 * tab, why it carries an identity.
 */
const IDENTITY_TABLE = `blip	r	schema	reference or id-shaped type in the Transitional schema graph
bottomLeft	r	schema	reference or id-shaped type in the Transitional schema graph
bottomRight	r	schema	reference or id-shaped type in the Transitional schema graph
cs	r	schema	reference or id-shaped type in the Transitional schema graph
dm	r	schema	reference or id-shaped type in the Transitional schema graph
durableId	*	extension	w15/w16cex annotation identity that survives a round-trip
embed	r	schema	reference or id-shaped type in the Transitional schema graph
href	r	schema	reference or id-shaped type in the Transitional schema graph
id	r,w	schema	reference or id-shaped type in the Transitional schema graph
Ignorable	*	extension	markup-compatibility processing directive; reading a foreign one keeps or drops the wrong AlternateContent branch
link	r	schema	reference or id-shaped type in the Transitional schema graph
lo	r	schema	reference or id-shaped type in the Transitional schema graph
name	w	extension	bookmark name and font-table name are cross-part join keys; the graph also declares a name attribute in the drawing and theme vocabularies
paraId	*	extension	w14/w15/w16cex paragraph identity; comments, revisions and annotations join to a paragraph through it
paraIdParent	*	extension	w15 comment-thread parent link
pict	r	schema	reference or id-shaped type in the Transitional schema graph
qs	r	schema	reference or id-shaped type in the Transitional schema graph
space	xml	extension	xml:space preserves significant whitespace in a run; the graph also declares w:cols/@w:space, which is column spacing
textId	*	extension	w14 paragraph text identity, paired with paraId across a revision
topLeft	r	schema	reference or id-shaped type in the Transitional schema graph
topRight	r	schema	reference or id-shaped type in the Transitional schema graph`;

const readIdentityTable = (): ReadonlyMap<string, IdentityAttribute> => {
  const identities = new Map<string, IdentityAttribute>();
  for (const line of IDENTITY_TABLE.split("\n")) {
    const [attribute, prefixes, source, reason] = line.split("\t");
    if (attribute === undefined || prefixes === undefined || reason === undefined) {
      continue;
    }
    identities.set(attribute, {
      prefixes: prefixes === "*" ? null : prefixes.split(","),
      reason,
      source: source === "extension" ? "extension" : "schema",
    });
  }
  return identities;
};

/** Attribute local name -> the identity it carries. */
export const IDENTITY_ATTRIBUTES: ReadonlyMap<string, IdentityAttribute> = readIdentityTable();
