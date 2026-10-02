/** Generate bounded runtime attribute facts; run with `write` or `check`. */
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import type { OoxmlSchemaGraph } from "./generate-ooxml-schema-graph";

class FuzzSchemaGenerationError extends TaggedError("FuzzSchemaGenerationError")<{
  message: string;
}> {}

const ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT = path.join(ROOT, "packages/docx-core/src/validate/schemaAttributes.gen.ts");
const XSD = "http://www.w3.org/2001/XMLSchema";
const WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type Whitespace = "preserve" | "replace" | "collapse";
type Constraint =
  | { type: "any"; whitespace: Whitespace }
  | { type: "enum"; values: string[]; whitespace: Whitespace }
  | { type: "pattern"; patterns: string[]; whitespace: Whitespace }
  | { type: "integer"; min?: string; max?: string; whitespace: Whitespace }
  | { type: "union"; members: Constraint[] };
type Attribute = { name: string; required: boolean; constraint: Constraint; fixed?: string };
type Model = { attributes: Attribute[]; children: Record<string, string[]> };

const INTEGER_BOUNDS: Record<string, { min?: string; max?: string }> = {
  integer: {},
  nonNegativeInteger: { min: "0" },
  positiveInteger: { min: "1" },
  nonPositiveInteger: { max: "0" },
  negativeInteger: { max: "-1" },
  byte: { min: "-128", max: "127" },
  short: { min: "-32768", max: "32767" },
  int: { min: "-2147483648", max: "2147483647" },
  long: { min: "-9223372036854775808", max: "9223372036854775807" },
  unsignedByte: { min: "0", max: "255" },
  unsignedShort: { min: "0", max: "65535" },
  unsignedInt: { min: "0", max: "4294967295" },
  unsignedLong: { min: "0", max: "18446744073709551615" },
};
const qualify = (namespace: string, name: string) => (namespace ? `{${namespace}}${name}` : name);

export const generateFuzzSchema = (graph: OoxmlSchemaGraph): string => {
  const symbols = new Map(graph.symbols.map((symbol) => [symbol.id, symbol]));
  const inheritance = new Map(graph.inheritance.map((edge) => [edge.derived, edge.base]));
  const attributes = Map.groupBy(graph.attributes, (attribute) => attribute.owner);
  const children = Map.groupBy(graph.children, (child) => child.owner);
  const whitespaceFor = (qname: string | undefined, visiting = new Set<string>()): Whitespace => {
    if (!qname || visiting.has(qname)) return "preserve";
    visiting.add(qname);
    if (qname.startsWith(`{${XSD}}`)) {
      const name = qname.slice(XSD.length + 2);
      if (name === "string" || name === "anySimpleType") return "preserve";
      return name === "normalizedString" ? "replace" : "collapse";
    }
    const symbol = symbols.get(`simpleType:${qname}`);
    const facet = symbol?.facets?.find(({ kind }) => kind === "whiteSpace");
    if (!facet) return whitespaceFor(symbol?.base, visiting);
    if (facet.value === "preserve" || facet.value === "replace" || facet.value === "collapse")
      return facet.value;
    throw new FuzzSchemaGenerationError({
      message: `Unknown whiteSpace facet on ${qname}: ${facet.value}`,
    });
  };
  const constraintFor = (qname: string | undefined, visiting = new Set<string>()): Constraint => {
    const whitespace = whitespaceFor(qname);
    if (!qname || visiting.has(qname)) return { type: "any", whitespace };
    visiting.add(qname);
    if (qname === `{${XSD}}boolean`)
      return { type: "enum", values: ["0", "1", "false", "true"], whitespace };
    const builtin = qname.startsWith(`{${XSD}}`)
      ? INTEGER_BOUNDS[qname.slice(XSD.length + 2)]
      : undefined;
    if (builtin) return { type: "integer", ...builtin, whitespace };
    const symbol = symbols.get(`simpleType:${qname}`);
    if (!symbol) return { type: "any", whitespace };
    if (symbol.enumValues)
      return { type: "enum", values: [...symbol.enumValues].sort(), whitespace };
    if (symbol.memberTypes)
      return {
        type: "union",
        members: symbol.memberTypes.map((member) => constraintFor(member, new Set(visiting))),
      };
    const base = constraintFor(symbol.base, visiting);
    const patterns = (symbol.facets ?? [])
      .filter((facet) => facet.kind === "pattern")
      .map((facet) => facet.value);
    // These WML lexical patterns use ordinary character classes, grouping,
    // quantifiers and escaped dots/braces; XSD-only regex syntax is excluded.
    if (patterns.length && patterns.every((pattern) => !/\\[ipPcCsSwW]|\[.*-\[/.test(pattern)))
      return { type: "pattern", patterns, whitespace };
    if (base.type !== "union") base.whitespace = whitespace;
    if (base.type !== "integer") return base;
    for (const facet of symbol.facets ?? []) {
      if (facet.kind === "minInclusive") base.min = facet.value;
      if (facet.kind === "maxInclusive") base.max = facet.value;
      if (facet.kind === "minExclusive") base.min = String(BigInt(facet.value) + 1n);
      if (facet.kind === "maxExclusive") base.max = String(BigInt(facet.value) - 1n);
    }
    return base;
  };
  const models: Record<string, Model> = {};
  const modelFor = (owner: string, visiting = new Set<string>()): Model => {
    const cached = models[owner];
    if (cached) return cached;
    if (visiting.has(owner)) return { attributes: [], children: {} };
    visiting.add(owner);
    const model: Model = { attributes: [], children: {} };
    const base = inheritance.get(owner);
    if (base) {
      const inherited = modelFor(`complexType:${base}`, new Set(visiting));
      model.attributes.push(...inherited.attributes);
      Object.assign(model.children, inherited.children);
    }
    for (const attribute of attributes.get(owner) ?? []) {
      if (attribute.kind === "group" && attribute.ref) {
        model.attributes.push(
          ...modelFor(`attributeGroup:${attribute.ref}`, new Set(visiting)).attributes,
        );
        continue;
      }
      if (attribute.kind !== "attribute") continue;
      const global = attribute.ref ? symbols.get(`attribute:${attribute.ref}`) : undefined;
      const name = attribute.ref ?? qualify(attribute.namespace ?? "", attribute.name ?? "");
      const constraint = attribute.enumValues
        ? {
            type: "enum" as const,
            values: [...attribute.enumValues].sort(),
            whitespace: whitespaceFor(attribute.type ?? global?.type),
          }
        : constraintFor(attribute.type ?? global?.type);
      // ECMA's xs:integer has no bound. Annotation interoperability uses a signed
      // 32-bit interoperability policy; this is deliberately not an XSD facet.
      if (name === `{${WML}}id` && constraint.type === "integer") {
        constraint.min = "-2147483648";
        constraint.max = "2147483647";
      }
      const result: Attribute = { name, required: attribute.use === "required", constraint };
      if (attribute.fixed !== undefined) result.fixed = attribute.fixed;
      const previous = model.attributes.findIndex((entry) => entry.name === name);
      if (previous >= 0) model.attributes.splice(previous, 1);
      if (attribute.use !== "prohibited") model.attributes.push(result);
    }
    for (const child of children.get(owner) ?? []) {
      if (child.kind === "group" && child.ref) {
        for (const [name, types] of Object.entries(
          modelFor(`group:${child.ref}`, new Set(visiting)).children,
        )) {
          const target = (model.children[name] ??= []);
          for (const type of types) if (!target.includes(type)) target.push(type);
        }
        continue;
      }
      if (child.kind !== "element") continue;
      const global = child.ref ? symbols.get(`element:${child.ref}`) : undefined;
      const type = child.type ?? global?.type;
      if (!type) continue;
      const name = child.ref ?? qualify(child.namespace ?? "", child.name ?? "");
      const target = (model.children[name] ??= []);
      if (!target.includes(type)) target.push(type);
    }
    model.attributes.sort((a, b) => a.name.localeCompare(b.name, "en"));
    models[owner] = model;
    return model;
  };
  const elements: Record<string, string[]> = {};
  for (const symbol of graph.symbols) {
    if (symbol.kind === "complexType") modelFor(symbol.id);
    if (symbol.kind !== "element" || !symbol.type) continue;
    elements[qualify(symbol.namespace, symbol.name)] = [symbol.type];
  }
  // Collect local WML declarations as fallback for wildcard payloads.
  for (const child of graph.children) {
    if (child.kind !== "element" || !child.name || !child.type) continue;
    const target = (elements[qualify(child.namespace ?? "", child.name)] ??= []);
    if (!target.includes(child.type)) target.push(child.type);
  }
  // WML owns story attributes; other part vocabularies receive XML syntax
  // checks in the caller. Keep contextual edges only for ambiguous WML names.
  const storyElements = Object.fromEntries(
    Object.entries(elements).filter(([name]) => name.startsWith(`{${WML}}`)),
  );
  const types = Object.fromEntries(
    Object.entries(models)
      .filter(([name]) => name.startsWith(`complexType:{${WML}}`))
      .map(([name, model]) => [
        name.slice("complexType:".length),
        {
          attributes: model.attributes.filter(
            (attribute) =>
              attribute.required ||
              attribute.fixed !== undefined ||
              attribute.constraint.type !== "any",
          ),
          children: Object.fromEntries(
            Object.entries(model.children).filter(
              ([child]) => (storyElements[child]?.length ?? 0) > 1,
            ),
          ),
        },
      ])
      .sort(([a], [b]) => String(a).localeCompare(String(b), "en")),
  );
  const namespaces = graph.namespaces.map(({ uri }) => uri).sort();
  const shorten = (value: string): string => {
    for (const [index, namespace] of namespaces.entries())
      value = value.replaceAll(`{${namespace}}`, `{${index}}`);
    return value;
  };
  let encoded = shorten(
    JSON.stringify({
      elements: Object.fromEntries(
        Object.entries(storyElements).sort(([a], [b]) => a.localeCompare(b, "en")),
      ),
      types,
    }),
  );
  for (const [index, name] of Object.keys(types).entries())
    encoded = encoded.replaceAll(JSON.stringify(shorten(name)), JSON.stringify(String(index)));
  const facts: { elements: Record<string, string[]>; types: Record<string, Model> } =
    JSON.parse(encoded);
  const definitions: Attribute[] = [];
  const definitionIndex = new Map<string, number>();
  for (const model of Object.values(facts.types)) {
    const indices = model.attributes.map((attribute) => {
      const key = JSON.stringify(attribute);
      const existing = definitionIndex.get(key);
      if (existing !== undefined) return existing;
      const index = definitions.length;
      definitions.push(attribute);
      definitionIndex.set(key, index);
      return index;
    });
    Object.assign(model, { attributes: indices });
  }
  return `// Generated by scripts/generate-fuzz-schema.ts; do not edit.\n// Attribute checks only, not complete XSD validation. WML id bounds are interoperability policy.\nimport type { SchemaAttributeFacts } from "./schemaAttributes";\n\nconst loadSchemaAttributeFacts = (json: string): SchemaAttributeFacts => JSON.parse(json);\n\nexport const SCHEMA_ATTRIBUTE_FACTS = loadSchemaAttributeFacts(\n  ${JSON.stringify(JSON.stringify({ namespaces, definitions, ...facts }))},\n);\n`;
};

if (import.meta.main) {
  const mode = process.argv.at(2);
  if (mode !== "write" && mode !== "check")
    throw new FuzzSchemaGenerationError({ message: "Expected write or check" });
  const graph: OoxmlSchemaGraph = JSON.parse(
    await readFile(
      path.join(ROOT, "specifications/generated/docx-transitional-schema.gen.json"),
      "utf8",
    ),
  );
  const output = generateFuzzSchema(graph);
  if (mode === "write") await writeFile(OUTPUT, output);
  else if ((await readFile(OUTPUT, "utf8")) !== output)
    throw new FuzzSchemaGenerationError({
      message: "Fuzz schema facts drifted; run bun scripts/generate-fuzz-schema.ts write",
    });
}
