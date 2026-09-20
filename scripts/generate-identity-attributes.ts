/**
 * Generate the set of OOXML attributes that carry an identity.
 *
 * `getAttribute(element, prefix, name)` resolves a prefixed attribute in three
 * steps and ends at an any-prefix local-name match, so a producer that binds
 * WordprocessingML to `ns0:` keeps its values. For a value that only names
 * itself that fallback is a tolerance; for a value that names something else -
 * a relationship target, a paragraph another part points at, a bookmark a
 * hyperlink anchors to - it is a way to read a foreign attribute as the real
 * one and join the wrong two things together.
 *
 * Which attributes those are is a property of the schema, not a list worth
 * keeping by hand: an attribute is identity-bearing when its type is a
 * reference type (`ST_RelationshipId`, `xsd:ID`/`IDREF`) or an id-shaped number
 * (`ST_DecimalNumber`, `ST_LongHexNumber` under an id-shaped name). The
 * derivation below reads them out of the Transitional schema graph.
 *
 * What the graph cannot supply is {@link EXTENSION_IDENTITY_ATTRIBUTES}: Word's
 * `w14`/`w15`/`w16cex` extension vocabularies are not part of the Transitional
 * schema set, and neither is markup compatibility or the `xml` namespace, yet
 * `w14:paraId` is the identity the comment, revision and annotation parts join
 * on. Each entry records why it is here and which prefixes it covers, and the
 * generator holds that scope to the graph: an entry claiming every prefix must
 * name a local name the graph declares nowhere, and an entry naming prefixes
 * must collide with one. A schema refresh that introduces a colliding name
 * fails here rather than widening the rule silently.
 *
 * Usage:
 *   bun scripts/generate-identity-attributes.ts write
 *   bun scripts/generate-identity-attributes.ts check
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import { loadSchemaGraph, localName, type SchemaGraph } from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "specifications/generated/identity-attributes.gen.ts");

class GenerateIdentityAttributesError extends TaggedError("GenerateIdentityAttributesError")<{
  message: string;
}> {}

const XSD_NAMESPACE = "http://www.w3.org/2001/XMLSchema";

/** Types whose value is a reference to something outside the element. */
const REFERENCE_TYPES = new Set([
  "ST_RelationshipId",
  `{${XSD_NAMESPACE}}ID`,
  `{${XSD_NAMESPACE}}IDREF`,
]);

/**
 * Numeric types that carry an identity only under an id-shaped name.
 *
 * `ST_DecimalNumber` is also every count and index in WordprocessingML, and
 * `ST_LongHexNumber` is also a color and a checksum, so the type alone decides
 * nothing; `w:id` on an annotation and `w14:paraId` on a paragraph are the
 * identities, and both are named for what they are.
 */
const ID_SHAPED_NAMES = new Set(["durableId", "id", "paraId", "paraIdParent", "textId"]);
const ID_SHAPED_NUMERIC_TYPES = new Set(["ST_DecimalNumber", "ST_LongHexNumber"]);

/**
 * The prefix each namespace is conventionally read through, or `null` when its
 * attributes are unqualified.
 *
 * The rule keys on the prefix a call names, so a namespace whose attributes
 * carry no prefix contributes nothing: `getAttribute`'s any-prefix fallback
 * only runs for a prefixed read, and a `.rels` part spells its `Id` bare.
 * The generator requires this table to be total over the namespaces its
 * derivation reaches.
 */
const NAMESPACE_PREFIXES: Readonly<Record<string, string | null>> = {
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
  "http://schemas.openxmlformats.org/package/2006/relationships": null,
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
};

type ExtensionEntry = {
  attribute: string;
  /** `null` claims every prefix: no other vocabulary declares this local name. */
  prefixes: readonly string[] | null;
  reason: string;
};

/** Identities no Transitional schema declares, and the scope each one claims. */
const EXTENSION_IDENTITY_ATTRIBUTES: readonly ExtensionEntry[] = [
  {
    attribute: "paraId",
    prefixes: null,
    reason:
      "w14/w15/w16cex paragraph identity; comments, revisions and annotations join to a paragraph through it",
  },
  {
    attribute: "textId",
    prefixes: null,
    reason: "w14 paragraph text identity, paired with paraId across a revision",
  },
  {
    attribute: "paraIdParent",
    prefixes: null,
    reason: "w15 comment-thread parent link",
  },
  {
    attribute: "durableId",
    prefixes: null,
    reason: "w15/w16cex annotation identity that survives a round-trip",
  },
  {
    attribute: "Ignorable",
    prefixes: null,
    reason:
      "markup-compatibility processing directive; reading a foreign one keeps or drops the wrong AlternateContent branch",
  },
  {
    attribute: "space",
    prefixes: ["xml"],
    reason:
      "xml:space preserves significant whitespace in a run; the graph also declares w:cols/@w:space, which is column spacing",
  },
  {
    attribute: "name",
    prefixes: ["w"],
    reason:
      "bookmark name and font-table name are cross-part join keys; the graph also declares a name attribute in the drawing and theme vocabularies",
  },
];

type IdentityAttribute = {
  attribute: string;
  prefixes: readonly string[] | null;
  reason: string;
  source: "extension" | "schema";
};

/** Every attribute local name the graph declares, however it is declared. */
const declaredLocalNames = (graph: SchemaGraph): ReadonlySet<string> => {
  const names = new Set<string>();
  for (const symbol of graph.symbols) {
    if (symbol.kind === "attribute") {
      names.add(symbol.name);
    }
  }
  for (const attribute of graph.attributes) {
    if (attribute.name !== undefined) {
      names.add(attribute.name);
    }
    if (attribute.ref !== undefined) {
      names.add(localName(attribute.ref));
    }
  }
  return names;
};

const isIdentityType = (name: string, type: string): boolean => {
  const local = localName(type);
  if (REFERENCE_TYPES.has(local) || REFERENCE_TYPES.has(type)) {
    return true;
  }
  return ID_SHAPED_NUMERIC_TYPES.has(local) && ID_SHAPED_NAMES.has(name);
};

type Declaration = { name: string; namespace: string };

/**
 * The identity-bearing attributes the graph declares, with the namespace each
 * one is read through.
 *
 * A `ref` names a global attribute and carries that attribute's namespace
 * (`r:id` on a WordprocessingML element stays in the relationship namespace);
 * an inline declaration belongs to its owner's namespace.
 */
const schemaDeclarations = (graph: SchemaGraph): Declaration[] => {
  const byId = new Map(graph.symbols.map((symbol) => [symbol.id, symbol]));
  const globals = new Map(
    graph.symbols
      .filter((symbol) => symbol.kind === "attribute")
      .map((symbol) => [`{${symbol.namespace}}${symbol.name}`, symbol]),
  );
  const declarations: Declaration[] = [];

  for (const symbol of globals.values()) {
    if (symbol.type !== undefined && isIdentityType(symbol.name, symbol.type)) {
      declarations.push({ name: symbol.name, namespace: symbol.namespace });
    }
  }

  for (const attribute of graph.attributes) {
    if (attribute.ref !== undefined) {
      const global = globals.get(attribute.ref);
      if (global?.type !== undefined && isIdentityType(global.name, global.type)) {
        declarations.push({ name: global.name, namespace: global.namespace });
      }
      continue;
    }
    if (attribute.name === undefined || attribute.type === undefined) {
      continue;
    }
    if (!isIdentityType(attribute.name, attribute.type)) {
      continue;
    }
    const owner = byId.get(attribute.owner);
    if (owner === undefined) {
      throw new GenerateIdentityAttributesError({
        message: `@${attribute.name} is declared on ${attribute.owner}, which the graph has no symbol for.`,
      });
    }
    declarations.push({ name: attribute.name, namespace: owner.namespace });
  }

  return declarations;
};

const collectIdentityAttributes = (graph: SchemaGraph): IdentityAttribute[] => {
  const prefixesByName = new Map<string, Set<string>>();
  for (const { name, namespace } of schemaDeclarations(graph)) {
    if (!(namespace in NAMESPACE_PREFIXES)) {
      throw new GenerateIdentityAttributesError({
        message: `@${name} is identity-bearing in ${namespace}, which NAMESPACE_PREFIXES does not name a prefix for. Add it, or record that its attributes are unqualified with null.`,
      });
    }
    const prefix = NAMESPACE_PREFIXES[namespace];
    if (prefix === null) {
      continue;
    }
    const prefixes = prefixesByName.get(name);
    if (prefixes === undefined) {
      prefixesByName.set(name, new Set([prefix]));
      continue;
    }
    prefixes.add(prefix);
  }

  const identities: IdentityAttribute[] = [...prefixesByName].map(([attribute, prefixes]) => ({
    attribute,
    prefixes: [...prefixes].toSorted(),
    reason: "reference or id-shaped type in the Transitional schema graph",
    source: "schema",
  }));

  const declared = declaredLocalNames(graph);
  for (const entry of EXTENSION_IDENTITY_ATTRIBUTES) {
    const collides = declared.has(entry.attribute);
    if (entry.prefixes === null && collides) {
      throw new GenerateIdentityAttributesError({
        message: `@${entry.attribute} claims every prefix, but the schema graph declares that local name too. Scope the entry to the prefixes it means.`,
      });
    }
    if (entry.prefixes !== null && !collides) {
      throw new GenerateIdentityAttributesError({
        message: `@${entry.attribute} is scoped to ${entry.prefixes.join(", ")}, but no other vocabulary declares that local name. Claim every prefix instead.`,
      });
    }
    if (identities.some((identity) => identity.attribute === entry.attribute)) {
      throw new GenerateIdentityAttributesError({
        message: `@${entry.attribute} is derived from the schema graph; drop the extension entry.`,
      });
    }
    identities.push({
      attribute: entry.attribute,
      prefixes: entry.prefixes,
      reason: entry.reason,
      source: "extension",
    });
  }

  return identities.toSorted((left, right) => left.attribute.localeCompare(right.attribute));
};

const renderModule = (identities: readonly IdentityAttribute[]): string => {
  const rows = identities.map(({ attribute, prefixes, reason, source }) => {
    if (/[`\t\n$\\]/u.test(reason)) {
      throw new GenerateIdentityAttributesError({
        message: `@${attribute}: a reason is rendered into a template literal and a tab-separated row, so it cannot carry a backtick, a dollar, a backslash, a tab or a newline.`,
      });
    }
    return `${attribute}\t${prefixes === null ? "*" : prefixes.join(",")}\t${source}\t${reason}`;
  });

  return `/**
 * GENERATED FILE - do not edit.
 *
 * The OOXML attributes whose value names something outside the element that
 * carries it, derived from the Transitional schema graph and Word's extension
 * vocabularies by \`scripts/generate-identity-attributes.ts\`. Regenerate with:
 *
 *   bun run generate:identity-attributes
 */

/** An attribute a prefix-resolved read can join to the wrong thing. */
export type IdentityAttribute = {
  /** The prefixes the identity is read through, or \`null\` for every prefix. */
  readonly prefixes: readonly string[] | null;
  readonly reason: string;
  readonly source: "extension" | "schema";
};

/**
 * One attribute per line: local name, a tab, the prefixes it is read through
 * (comma-separated, or \`*\` for every prefix), a tab, where it came from, a
 * tab, why it carries an identity.
 */
const IDENTITY_TABLE = \`${rows.join("\n")}\`;

const readIdentityTable = (): ReadonlyMap<string, IdentityAttribute> => {
  const identities = new Map<string, IdentityAttribute>();
  for (const line of IDENTITY_TABLE.split("\\n")) {
    const [attribute, prefixes, source, reason] = line.split("\\t");
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
`;
};

const digest = (value: string): string => createHash("sha256").update(value).digest("hex");

const main = async (): Promise<void> => {
  const mode = process.argv.at(2) ?? "write";
  if (mode !== "check" && mode !== "write") {
    throw new GenerateIdentityAttributesError({
      message: "Usage: bun scripts/generate-identity-attributes.ts [check|write]",
    });
  }

  const graph = await loadSchemaGraph();
  const identities = collectIdentityAttributes(graph);

  const relationshipId = identities.find(({ attribute }) => attribute === "id");
  if (relationshipId?.prefixes?.includes("r") !== true) {
    throw new GenerateIdentityAttributesError({
      message: "r:id must be an identity attribute; the schema graph or the derivation drifted.",
    });
  }

  const rendered = renderModule(identities);
  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(
      `identity-attributes.gen.ts written (${identities.length} attributes: ` +
        `${identities.filter(({ source }) => source === "schema").length} from the schema graph, ` +
        `${identities.filter(({ source }) => source === "extension").length} from the extension vocabularies)\n`,
    );
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing === null || digest(existing) !== digest(rendered)) {
    throw new GenerateIdentityAttributesError({
      message: "identity-attributes.gen.ts is stale. Run `bun run generate:identity-attributes`.",
    });
  }
  process.stdout.write("identity-attributes.gen.ts is up to date\n");
};

await main();
