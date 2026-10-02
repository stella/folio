import { panic, Result } from "better-result";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { createXmlEntityDecoder } from "./xmlEntityDecoder";

import { SCHEMA_ATTRIBUTE_FACTS } from "./schemaAttributes.gen";

type Whitespace = "preserve" | "replace" | "collapse";
type Constraint =
  | { type: "any"; whitespace: Whitespace }
  | { type: "enum"; values: readonly string[]; whitespace: Whitespace }
  | { type: "pattern"; patterns: readonly string[]; whitespace: Whitespace }
  | { type: "integer"; min?: string; max?: string; whitespace: Whitespace }
  | { type: "union"; members: readonly Constraint[] };
type Attribute = { name: string; required: boolean; constraint: Constraint; fixed?: string };
type Model = {
  attributes: readonly number[];
  children: Readonly<Record<string, readonly string[]>>;
};
/** A record annotation bounds generated type inference and permits dynamic lookup. */
export type SchemaAttributeFacts = {
  namespaces: readonly string[];
  definitions: readonly Attribute[];
  elements: Readonly<Record<string, readonly string[]>>;
  types: Readonly<Record<string, Model>>;
};

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const namespaceIndex = new Map(SCHEMA_ATTRIBUTE_FACTS.namespaces.map((uri, index) => [uri, index]));
const strictToTransitional = (uri: string): string =>
  uri
    .replace(
      "http://purl.oclc.org/ooxml/wordprocessingml/main",
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    )
    .replace(
      "http://purl.oclc.org/ooxml/drawingml/",
      "http://schemas.openxmlformats.org/drawingml/2006/",
    )
    .replace(
      "http://purl.oclc.org/ooxml/officeDocument/",
      "http://schemas.openxmlformats.org/officeDocument/2006/",
    );

type ResolveNameOptions = { name: string; scope: ReadonlyMap<string, string>; attribute: boolean };
const resolveName = ({ name, scope, attribute }: ResolveNameOptions): string | null => {
  const colon = name.indexOf(":");
  const prefix = colon < 0 ? "" : name.slice(0, colon);
  const local = colon < 0 ? name : name.slice(colon + 1);
  const uri = prefix === "" && attribute ? "" : scope.get(prefix);
  if (uri === undefined && prefix !== "") return null;
  if (!uri) return local;
  const canonical = strictToTransitional(uri);
  return `{${namespaceIndex.get(canonical) ?? canonical}}${local}`;
};

/** XSD whiteSpace recognizes XML's four ASCII whitespace characters only. */
const normalizeWhitespace = (value: string, policy: Whitespace): string => {
  switch (policy) {
    case "preserve":
      return value;
    case "replace":
      return value.replace(/[\t\n\r]/g, " ");
    case "collapse":
      return value.replace(/[\t\n\r ]+/g, " ").replace(/^ | $/g, "");
    default: {
      const exhaustive: never = policy;
      return exhaustive;
    }
  }
};

const accepts = (constraint: Constraint, value: string): boolean => {
  const normalized =
    constraint.type === "union" ? value : normalizeWhitespace(value, constraint.whitespace);
  switch (constraint.type) {
    case "any":
      return true;
    case "enum":
      return constraint.values.includes(normalized);
    case "pattern":
      return constraint.patterns.every((pattern) =>
        new RegExp(`^(?:${pattern})$`, "u").test(normalized),
      );
    case "union":
      return constraint.members.some((member) => accepts(member, value));
    case "integer": {
      if (!/^[+-]?\d+$/.test(normalized)) return false;
      const integer = BigInt(normalized);
      return (
        (constraint.min === undefined || integer >= BigInt(constraint.min)) &&
        (constraint.max === undefined || integer <= BigInt(constraint.max))
      );
    }
    default: {
      const exhaustive: never = constraint;
      return exhaustive;
    }
  }
};

const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const validateAttributes = (model: Model, values: ReadonlyMap<string, string>): string | null => {
  for (const index of model.attributes) {
    const attribute = SCHEMA_ATTRIBUTE_FACTS.definitions.at(index);
    if (!attribute) return panic("Invalid generated schema attribute index");
    const value = values.get(attribute.name);
    if (value === undefined) {
      if (attribute.required) return `Missing required attribute ${attribute.name}`;
      continue;
    }
    if (attribute.fixed !== undefined && value !== attribute.fixed)
      return `Invalid fixed attribute ${attribute.name}`;
    if (!accepts(attribute.constraint, value))
      return `Invalid attribute value for ${attribute.name}`;
  }
  return null;
};

/**
 * Attribute-only oracle, not complete XSD validation. The caller bounds XML size.
 * Checks inherited required attributes, enum and integer values; signed 32-bit
 * WML id bounds are an interoperability policy (the XSD integer is unbounded).
 * Unknown extension vocabularies remain available for lossless round trips.
 */
export const validateSchemaAttributes = (xml: string): string | null => {
  const syntax = XMLValidator.validate(xml);
  if (syntax !== true) return `Malformed XML: ${syntax.err.msg}`;
  if (/<!DOCTYPE/i.test(xml)) return "DOCTYPE is not allowed in DOCX XML";
  const parsed = Result.try(() =>
    new XMLParser({
      preserveOrder: true,
      ignoreAttributes: false,
      attributeNamePrefix: "",
      parseTagValue: false,
      parseAttributeValue: false,
      trimValues: false,
      processEntities: true,
      entityDecoder: createXmlEntityDecoder(),
      ignoreDeclaration: true,
    }).parse(xml),
  );
  if (parsed.isErr()) return "Malformed XML: XML parser rejected the part";
  if (!Array.isArray(parsed.value)) return "Malformed XML: expected a document element";
  const roots = parsed.value.filter(
    (node: unknown) =>
      record(node) &&
      Object.keys(node).some((tag) => tag !== ":@" && !tag.startsWith("#") && !tag.startsWith("?")),
  );
  if (roots.length !== 1) return "Malformed XML: expected exactly one document element";
  const pending: Array<{
    nodes: unknown;
    scope: ReadonlyMap<string, string>;
    types: readonly string[];
  }> = [{ nodes: parsed.value, scope: new Map([["xml", XML_NAMESPACE]]), types: [] }];
  while (pending.length > 0) {
    const current = pending.pop();
    if (!current || !Array.isArray(current.nodes)) continue;
    for (const node of current.nodes) {
      if (!record(node)) continue;
      const attributes = record(node[":@"]) ? node[":@"] : {};
      const scope = new Map(current.scope);
      for (const [name, value] of Object.entries(attributes)) {
        if (typeof value !== "string") continue;
        if (name === "xmlns") scope.set("", value);
        else if (name.startsWith("xmlns:")) scope.set(name.slice(6), value);
      }
      const values = new Map<string, string>();
      for (const [name, value] of Object.entries(attributes)) {
        if (name === "xmlns" || name.startsWith("xmlns:")) continue;
        const resolved = resolveName({ name, scope, attribute: true });
        if (resolved === null) return `Unbound XML attribute prefix: ${name}`;
        if (values.has(resolved)) return `Duplicate XML attribute: ${name}`;
        if (typeof value === "string") values.set(resolved, value);
      }
      for (const [tag, nodes] of Object.entries(node)) {
        if (tag === ":@" || tag.startsWith("#") || tag.startsWith("?")) continue;
        const name = resolveName({ name: tag, scope, attribute: false });
        if (name === null) return `Unbound XML element prefix: ${tag}`;
        const contextual = current.types.flatMap(
          (type) => SCHEMA_ATTRIBUTE_FACTS.types[type]?.children[name] ?? [],
        );
        const types = contextual.length
          ? contextual
          : (SCHEMA_ATTRIBUTE_FACTS.elements[name] ?? []);
        let issue: string | null = null;
        for (const type of types) {
          const model = SCHEMA_ATTRIBUTE_FACTS.types[type];
          if (!model) continue;
          issue = validateAttributes(model, values);
          if (issue === null) break;
        }
        if (issue !== null) return `${tag}: ${issue}`;
        pending.push({ nodes, scope, types });
      }
    }
  }
  return null;
};
