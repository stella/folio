import { Result, TaggedError } from "better-result";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { createXmlEntityDecoder } from "./xmlEntityDecoder";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const WORDPROCESSING_PART_ROLES = new Set([
  "comments",
  "footnotes",
  "endnotes",
  "header",
  "footer",
  "styles",
  "numbering",
  "settings",
  "fontTable",
  "webSettings",
]);
const OFFICE_DOCUMENT_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
]);
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  entityDecoder: createXmlEntityDecoder(),
  ignoreDeclaration: true,
});
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const localName = (name: string) => name.slice(name.indexOf(":") + 1);
const namespace = (name: string, scope: Record<string, unknown>) => {
  const separator = name.indexOf(":");
  return scope[separator < 0 ? "xmlns" : `xmlns:${name.slice(0, separator)}`];
};
type PackageElement = {
  name: string;
  namespace: unknown;
  scope: Record<string, unknown>;
  attributes: Record<string, unknown>;
};
const elements = (xml: string): PackageElement[] | null => {
  if (XMLValidator.validate(xml) !== true) return null;
  const tree: unknown = parser.parse(xml);
  if (!Array.isArray(tree)) return null;
  const output: PackageElement[] = [];
  const visit = (nodes: unknown[], parent: Record<string, unknown>) => {
    for (const node of nodes) {
      if (!isRecord(node)) continue;
      const name = Object.keys(node).find(
        (key) => key !== ":@" && !key.startsWith("#") && !key.startsWith("?"),
      );
      if (!name) continue;
      const attributes = isRecord(node[":@"]) ? node[":@"] : {};
      const scope = Object.assign({}, parent, attributes);
      output.push({ name: localName(name), namespace: namespace(name, scope), scope, attributes });
      const children = node[name];
      if (Array.isArray(children)) visit(children, scope);
    }
  };
  visit(tree, {});
  return output;
};
const relationshipSource = (path: string): string | null => {
  if (path === "_rels/.rels") return "";
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/u.exec(path);
  return match ? `${match.at(1) ?? ""}${match.at(2) ?? ""}` : null;
};
class PackageUriError extends TaggedError("PackageUriError")<{ message: string }> {}

const resolveTarget = (source: string, target: string): string | null => {
  // OPC part URI resolution uses the source part's directory, never the .rels directory.
  const targetPath = target.split("#").at(0);
  if (
    !targetPath ||
    targetPath.includes("\\") ||
    targetPath.includes("?") ||
    /^[a-z]+:/iu.test(targetPath)
  )
    return null;
  const segments = targetPath.startsWith("/") ? [] : source.split("/").slice(0, -1);
  for (const segment of targetPath.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      if (segments.length === 0) return null;
      segments.pop();
    } else {
      const decoded = Result.try({
        try: () => decodeURIComponent(segment),
        catch: () => new PackageUriError({ message: "Malformed part URI" }),
      });
      if (
        decoded.isErr() ||
        decoded.value.includes("/") ||
        decoded.value.includes("\\") ||
        decoded.value === "." ||
        decoded.value === ".."
      )
        return null;
      segments.push(decoded.value);
    }
  }
  return segments.join("/");
};

/** OPC presence and relationship graph checks over already bounded, inflated XML parts. */
export const validatePackageGraph = (
  parts: ReadonlyMap<string, string>,
  paths: ReadonlySet<string>,
): string | null => {
  const contentTypes = elements(parts.get("[Content_Types].xml") ?? "");
  const root = contentTypes?.at(0);
  if (root?.name !== "Types" || root.namespace !== CONTENT_TYPES_NS)
    return "Invalid OPC content-types root";
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const element of contentTypes ?? []) {
    if (element === root) continue;
    if (element.namespace !== CONTENT_TYPES_NS) return "Invalid content-type declaration namespace";
    const type = element.attributes["ContentType"];
    if (typeof type !== "string" || type.length === 0) return "Missing content type";
    if (element.name === "Default") {
      const extension = element.attributes["Extension"];
      if (typeof extension !== "string" || !extension || defaults.has(extension.toLowerCase()))
        return "Invalid or duplicate default content type";
      defaults.set(extension.toLowerCase(), type);
    } else if (element.name === "Override") {
      const name = element.attributes["PartName"];
      if (typeof name !== "string" || !name.startsWith("/"))
        return "Invalid content-type part name";
      const resolved = resolveTarget("", name);
      if (resolved === null || overrides.has(resolved) || !paths.has(resolved))
        return "Invalid or dangling content-type override";
      overrides.set(resolved, type);
    } else return "Unknown content-type declaration";
  }
  const documentType = overrides.get("word/document.xml") ?? defaults.get("xml");
  if (
    documentType !==
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
  )
    return "Invalid officeDocument content type";
  for (const path of paths) {
    if (path === "[Content_Types].xml") continue;
    const segment = path.slice(path.lastIndexOf("/") + 1);
    const dot = segment.lastIndexOf(".");
    const extension = dot < 0 ? undefined : segment.slice(dot + 1).toLowerCase();
    if (!overrides.has(path) && (extension === undefined || !defaults.has(extension)))
      return `Missing content type for ${path}`;
  }
  let documentRelationship = false;
  const partRelationshipIds = new Map<string, Set<string>>();
  for (const [path, xml] of parts) {
    if (!path.endsWith(".rels")) continue;
    const source = relationshipSource(path);
    if (source === null || (source !== "" && !paths.has(source)))
      return `Dangling relationship source: ${path}`;
    const relationships = elements(xml);
    const relationshipRoot = relationships?.at(0);
    if (
      relationshipRoot?.name !== "Relationships" ||
      relationshipRoot.namespace !== RELATIONSHIPS_NS
    )
      return `Invalid relationships root: ${path}`;
    const ids = new Set<string>();
    partRelationshipIds.set(source, ids);
    for (const element of relationships ?? []) {
      if (element === relationshipRoot) continue;
      if (element.name !== "Relationship" || element.namespace !== RELATIONSHIPS_NS)
        return `Invalid relationship element: ${path}`;
      const { Id: id, Type: type, Target: target, TargetMode: mode } = element.attributes;
      if (
        typeof id !== "string" ||
        !id ||
        ids.has(id) ||
        typeof type !== "string" ||
        !type ||
        typeof target !== "string" ||
        !target
      )
        return `Invalid or duplicate relationship: ${path}`;
      ids.add(id);
      if (mode !== undefined && mode !== "Internal" && mode !== "External")
        return `Invalid relationship TargetMode: ${path}`;
      if (mode === "External") continue;
      const resolved = resolveTarget(source, target);
      if (resolved === null || !paths.has(resolved))
        return `Missing relationship target: ${path} -> ${target}`;
      const role = type.slice(type.lastIndexOf("/") + 1);
      const officeRelationship =
        type.startsWith("http://schemas.openxmlformats.org/officeDocument/2006/relationships/") ||
        type.startsWith("http://purl.oclc.org/ooxml/officeDocument/relationships/");
      if (officeRelationship && WORDPROCESSING_PART_ROLES.has(role)) {
        const declaredType =
          overrides.get(resolved) ??
          defaults.get(resolved.slice(resolved.lastIndexOf(".") + 1).toLowerCase());
        if (
          declaredType !==
          `application/vnd.openxmlformats-officedocument.wordprocessingml.${role}+xml`
        )
          return `Invalid content type for ${role}: ${resolved}`;
      }
      if (source === "" && OFFICE_DOCUMENT_TYPES.has(type) && resolved === "word/document.xml")
        documentRelationship = true;
    }
  }
  const relationshipNamespaces = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
  ]);
  for (const [path, xml] of parts) {
    if (path.endsWith(".rels") || path === "[Content_Types].xml") continue;
    const nodes = elements(xml);
    if (nodes === null) return `Malformed XML: ${path}`;
    for (const node of nodes) {
      for (const [name, value] of Object.entries(node.attributes)) {
        if (!name.includes(":")) continue;
        if (!relationshipNamespaces.has(String(namespace(name, node.scope)))) continue;
        if (typeof value !== "string" || !partRelationshipIds.get(path)?.has(value))
          return `Unresolved relationship attribute: ${path} ${name}=${String(value)}`;
      }
    }
  }
  return documentRelationship ? null : "Missing package officeDocument relationship";
};
