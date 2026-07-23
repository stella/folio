/**
 * XML Parser Utilities for OOXML
 *
 * Provides helper functions for parsing Office Open XML (OOXML) content
 * with proper namespace handling.
 *
 * OOXML uses many namespaces:
 * - w:  WordprocessingML (main document content)
 * - a:  DrawingML (graphics)
 * - r:  Relationships
 * - wp: Word Drawing positioning
 * - wps: Word Drawing shapes
 * - wpc: Word Drawing canvas
 * - wpg: Word Drawing group
 * - m:  Math
 * - mc: Markup Compatibility
 * - v:  VML (legacy vector graphics)
 * - o:  Office (extensions)
 * - pic: Pictures
 */

import { XMLBuilder, XMLParser } from "fast-xml-parser";

import { OOXML_NS } from "@stll/docx-utils";

/**
 * XML element tree node — drop-in replacement for the `Element` type
 * previously imported from `xml-js`. Every consumer imports this from
 * `xmlParser.ts`, so the shape must stay identical.
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
  elements?: XmlElement[];
};

// ---------------------------------------------------------------------------
// fast-xml-parser instances (reused across calls)
// ---------------------------------------------------------------------------

const fxpParserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false,
  textNodeName: "#text",
  parseTagValue: false,
  parseAttributeValue: false,
  ignoreDeclaration: true,
  ignorePiTags: true,
  // Folio leaves the parser callbacks at their defaults, which do not consume
  // the current-node path. Re-enable this if a callback needs a string path.
  jPath: false,
  // Security: only process the 5 built-in XML entities (&lt; &gt; &amp;
  // &apos; &quot;). Custom/DOCTYPE-defined entities are blocked to
  // prevent Billion Laughs exponential expansion DoS. OOXML does not
  // use custom entities.
  processEntities: true,
  htmlEntities: true,
};

const fxpParserOptionsWithStopNodes = {
  ...fxpParserOptions,
  // Skip parsing large base64 blobs into memory early. Enabled only for
  // XML parts that actually contain legacy inline binary payloads because
  // fast-xml-parser's stop-node matcher runs on every tag.
  stopNodes: ["*.w:binData"],
};

const fxpBuilderOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  textNodeName: "#text",
  suppressEmptyNode: true,
};

const fxpParser = new XMLParser(fxpParserOptions);
const fxpParserWithStopNodes = new XMLParser(fxpParserOptionsWithStopNodes);
const fxpBuilder = new XMLBuilder(fxpBuilderOptions);

// ---------------------------------------------------------------------------
// Converters: fast-xml-parser preserveOrder <-> XmlElement
// ---------------------------------------------------------------------------

/** Text node key used by fast-xml-parser in preserveOrder mode. */
const TEXT_KEY = "#text";
/** Attribute group key used by fast-xml-parser in preserveOrder mode. */
const ATTR_KEY = ":@";
/** Character reference required to keep carriage returns through XML end-of-line normalization. */
const XML_CARRIAGE_RETURN_REFERENCE = "&#13;";

type MutableFxpNode = XmlElement & Record<string, unknown>;

/**
 * Convert a fast-xml-parser preserveOrder node into an XmlElement.
 *
 * In preserveOrder mode every node is an object with exactly one "real" key
 * (the tag name or `#text`) whose value is the children array, plus an
 * optional `:@` key holding the attributes object.
 */
function fxpNodeToElement(node: Record<string, unknown>): MutableFxpNode {
  // Text node: { "#text": "some text" }
  if (TEXT_KEY in node) {
    return { type: "text", text: node[TEXT_KEY] as string };
  }

  // Element node: { "w:p": [...children], ":@": { ...attrs } }
  const attrs = node[ATTR_KEY] as Record<string, string> | undefined;

  // fast-xml-parser creates these nodes and rejects prototype-polluting tag
  // names. Iterating directly avoids allocating a keys array for every node.
  for (const key in node) {
    if (key === ATTR_KEY) {
      continue;
    }

    const children = node[key] as MutableFxpNode[];
    const element: MutableFxpNode = { type: "element", name: key };

    // fast-xml-parser only emits the attribute group when it contains an
    // attribute, so a second Object.keys allocation is unnecessary.
    if (attrs) {
      element.attributes = attrs;
    }

    if (children.length > 0) {
      for (let index = 0; index < children.length; index += 1) {
        children[index] = fxpNodeToElement(children[index]!);
      }
      element.elements = children;
    }

    return element;
  }

  // Shouldn't happen, but return an empty element as fallback
  return { type: "element" };
}

/**
 * Convert the top-level fast-xml-parser preserveOrder array into an
 * XmlElement that matches xml-js's non-compact root structure:
 * `{ elements: [...] }`.
 */
function fxpToRootElement(nodes: MutableFxpNode[]): XmlElement {
  for (let index = 0; index < nodes.length; index += 1) {
    nodes[index] = fxpNodeToElement(nodes[index]!);
  }

  return {
    elements: nodes,
  };
}

/**
 * Convert an XmlElement back into the fast-xml-parser preserveOrder format
 * so we can feed it to XMLBuilder.
 */
function elementToFxpNode(el: XmlElement): Record<string, unknown> {
  if (el.type === "text") {
    return { [TEXT_KEY]: el.text ?? "" };
  }

  const name = el.name ?? "";
  const children: Record<string, unknown>[] = el.elements ? el.elements.map(elementToFxpNode) : [];

  const node: Record<string, unknown> = { [name]: children };

  if (el.attributes && Object.keys(el.attributes).length > 0) {
    node[ATTR_KEY] = el.attributes;
  }

  return node;
}

/**
 * Common OOXML namespace URIs — re-exported from @stll/docx-utils.
 */
export const NAMESPACES = OOXML_NS;

/**
 * Parse XML string into element tree
 *
 * @param xml - XML string to parse
 * @returns Parsed element tree
 */
export function parseXml(xml: string): XmlElement {
  // fast-xml-parser with preserveOrder returns an array of nodes.
  // We convert it into the same tree shape that xml-js used to produce
  // (non-compact mode with attributesKey="attributes", textKey="text").
  //
  // IMPORTANT: trimValues is false so whitespace-only text nodes such as
  // <w:t xml:space="preserve"> </w:t> are preserved — matching the old
  // xml-js captureSpacesBetweenElements behaviour.
  const parser = xml.includes("binData") ? fxpParserWithStopNodes : fxpParser;
  const nodes = parser.parse(xml) as MutableFxpNode[];
  return fxpToRootElement(nodes);
}

/**
 * Serialize an XmlElement back to an XML string
 */
export function elementToXml(element: XmlElement): string {
  const fxpNode = elementToFxpNode(element);
  const xml = fxpBuilder.build([fxpNode]) as string;
  return xml.replaceAll("\r", XML_CARRIAGE_RETURN_REFERENCE);
}

/**
 * Parse XML string to a more convenient format
 */
export function parseXmlDocument(xml: string): XmlElement | null {
  try {
    const parsed = parseXml(xml);

    // The root is typically the declaration + elements array
    if (parsed.elements && parsed.elements.length > 0) {
      // Return the first real element (skip declarations)
      for (const element of parsed.elements) {
        if (element.type === "element") {
          return element;
        }
      }
      return null;
    }

    return parsed;
  } catch {
    return null;
  }
}

/**
 * Get local name from a prefixed element name
 * e.g., "w:p" -> "p", "a:graphic" -> "graphic"
 */
export function getLocalName(name: string | undefined): string {
  if (!name) {
    return "";
  }
  if (name[1] === ":") {
    return name.slice(2);
  }
  const colonIndex = name.indexOf(":");
  return colonIndex !== -1 ? name.slice(colonIndex + 1) : name;
}

/**
 * Get namespace prefix from an element name
 * e.g., "w:p" -> "w", "a:graphic" -> "a"
 */
export function getNamespacePrefix(name: string): string | null {
  if (name[1] === ":") {
    return name[0] ?? null;
  }
  const colonIndex = name.indexOf(":");
  return colonIndex !== -1 ? name.slice(0, colonIndex) : null;
}

function hasLocalName(name: string | undefined, localName: string): boolean {
  if (!name) {
    return false;
  }
  if (name === localName) {
    return true;
  }
  const localNameOffset = name.length - localName.length;
  return localNameOffset > 0 && name[localNameOffset - 1] === ":" && name.endsWith(localName);
}

/**
 * Check if an element matches a given namespaced name
 *
 * @param element - Element to check
 * @param namespace - Namespace prefix (e.g., "w", "a")
 * @param localName - Local element name (e.g., "p", "r")
 */
export function matchesName(element: XmlElement, _namespace: string, localName: string): boolean {
  if (!element.name) {
    return false;
  }

  // Namespace prefixes are aliases, so preserve the existing any-prefix
  // compatibility without allocating a canonical qualified name.
  return hasLocalName(element.name, localName);
}

/**
 * Find first child element matching the given namespaced name
 *
 * @param parent - Parent element
 * @param namespace - Namespace prefix (e.g., "w")
 * @param localName - Local element name (e.g., "p")
 * @returns First matching child or null
 */
export function findChild(
  parent: XmlElement | null | undefined,
  _namespace: string,
  localName: string,
): XmlElement | null {
  if (!parent || !parent.elements) {
    return null;
  }

  for (const child of parent.elements) {
    if (child.type !== "element") {
      continue;
    }

    // Namespace prefixes are aliases, so this recognizes the canonical,
    // alternate-prefix, and unprefixed forms in one allocation-free check.
    if (hasLocalName(child.name, localName)) {
      return child;
    }
  }

  return null;
}

/**
 * Find all child elements matching the given namespaced name
 *
 * @param parent - Parent element
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns Array of matching children
 */
export function findChildren(
  parent: XmlElement | null | undefined,
  _namespace: string,
  localName: string,
): XmlElement[] {
  if (!parent || !parent.elements) {
    return [];
  }

  const results: XmlElement[] = [];

  for (const child of parent.elements) {
    if (child.type !== "element") {
      continue;
    }

    if (hasLocalName(child.name, localName)) {
      results.push(child);
    }
  }

  return results;
}

/**
 * Find first child element by local name only (ignoring namespace)
 *
 * @param parent - Parent element
 * @param localName - Local element name
 * @returns First matching child or null
 */
export function findChildByLocalName(
  parent: XmlElement | null | undefined,
  localName: string,
): XmlElement | null {
  if (!parent || !parent.elements) {
    return null;
  }

  for (const child of parent.elements) {
    if (child.type !== "element") {
      continue;
    }

    if (hasLocalName(child.name, localName)) {
      return child;
    }
  }

  return null;
}

/**
 * Find all child elements by local name only
 *
 * @param parent - Parent element
 * @param localName - Local element name
 * @returns Array of matching children
 */
export function findChildrenByLocalName(
  parent: XmlElement | null | undefined,
  localName: string,
): XmlElement[] {
  if (!parent || !parent.elements) {
    return [];
  }

  const results: XmlElement[] = [];
  for (const child of parent.elements) {
    if (child.type === "element" && hasLocalName(child.name, localName)) {
      results.push(child);
    }
  }
  return results;
}

/**
 * Find first child element by full name (including namespace prefix)
 *
 * @param parent - Parent element
 * @param fullName - Full element name with namespace prefix (e.g., 'wp:extent')
 * @returns First matching child or null
 */
export function findByFullName(
  parent: XmlElement | null | undefined,
  fullName: string,
): XmlElement | null {
  if (!parent || !parent.elements) {
    return null;
  }

  for (const child of parent.elements) {
    if (child.type !== "element") {
      continue;
    }
    if (child.name === fullName) {
      return child;
    }
  }

  return null;
}

/**
 * Get all child elements (excludes text nodes, etc.)
 *
 * @param parent - Parent element
 * @returns Array of child elements
 */
export function getChildElements(parent: XmlElement | null | undefined): XmlElement[] {
  if (!parent || !parent.elements) {
    return [];
  }
  const results: XmlElement[] = [];
  for (const child of parent.elements) {
    if (child.type === "element") {
      results.push(child);
    }
  }
  return results;
}

/**
 * Get an attribute value from an element
 *
 * @param element - Element to get attribute from
 * @param namespace - Namespace prefix for the attribute (or null for no namespace)
 * @param name - Attribute name
 * @returns Attribute value or null if not found
 *
 * Lookup order when a namespace is provided:
 *   1. exact `${namespace}:${name}` match (`w:val`),
 *   2. bare `${name}` match (`val`),
 *   3. any-prefix local-name match (`x:val`, `w14:val`, …).
 *
 * The local-name fallback mirrors `findChild` and lets the parser tolerate
 * producers that rebind the WordprocessingML namespace to an alternative
 * prefix. Without it, attributes such as `x:val` on an `x:alias` element
 * would silently drop on round-trip even though the element itself is
 * recognised.
 */
export function getAttribute(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string,
): string | null {
  if (!element || !element.attributes) {
    return null;
  }

  const attrs = element.attributes as Record<string, string>;

  // Try with namespace prefix first
  if (namespace) {
    const prefixedValue = attrs[`${namespace}:${name}`];
    if (prefixedValue !== undefined) {
      return prefixedValue;
    }
  }

  // Try without namespace
  const unprefixedValue = attrs[name];
  if (unprefixedValue !== undefined) {
    return unprefixedValue;
  }

  // Fall back to any-prefix local-name match so alt-prefix producers keep
  // their attribute values. Only applies when the caller asked for a
  // namespaced attribute; if `namespace` is null the caller explicitly
  // wanted an unprefixed attribute.
  if (namespace) {
    for (const key in attrs) {
      if (hasLocalName(key, name)) {
        return attrs[key] ?? null;
      }
    }
  }

  return null;
}

/**
 * Read an attribute by local name, trying the element's own prefix first,
 * then the canonical Word prefix, then unprefixed.
 *
 * OOXML lets a producer bind the WordprocessingML namespace under any
 * prefix the file's xmlns declarations choose, so a strict `w:val` /
 * `w:tag` lookup misses valid docs that use `ns0:` (or any other prefix
 * resolving to the same URI). This helper mirrors `parseBooleanElement`'s
 * tolerance for the attribute path — call it whenever you want to read
 * an OOXML attribute that conventionally lives in the `w:` namespace
 * but could appear under an alternate prefix bound to the same URI.
 */
export function getAttributeAnyPrefix(
  element: XmlElement | null | undefined,
  localName: string,
): string | null {
  if (!element) {
    return null;
  }
  const elementName = element.name ?? "";
  const colonIdx = elementName.indexOf(":");
  const elementPrefix = colonIdx > 0 ? elementName.slice(0, colonIdx) : null;
  if (elementPrefix && elementPrefix !== "w") {
    const fromPrefix = getAttribute(element, elementPrefix, localName);
    if (fromPrefix !== null) {
      return fromPrefix;
    }
  }
  const canonical = getAttribute(element, "w", localName);
  if (canonical !== null) {
    return canonical;
  }
  // Final fallback: a canonical element (`<w:tag>`) can still carry
  // attributes under a different inherited prefix (`<w:tag x:val="…"/>`
  // when `x` is also bound to the WP URI at the document root). The
  // wrapper-prefix and canonical-prefix lookups above miss that case;
  // scan attribute names by local-name suffix as a last resort so the
  // modeled projection picks the value up — without this, the raw XML
  // normalizer rewrites it on save but `props.tag` / `listItems` /
  // similar stay empty and tag-keyed lookups break.
  const attrs = element.attributes;
  if (attrs) {
    for (const key in attrs) {
      if (hasLocalName(key, localName)) {
        const value = attrs[key];
        if (value === undefined) {
          return null;
        }
        return typeof value === "string" ? value : String(value);
      }
    }
  }
  return null;
}

/**
 * Get an attribute value, trying multiple possible names
 *
 * @param element - Element to get attribute from
 * @param names - Array of possible attribute names (with or without namespace)
 * @returns First found attribute value or null
 */
export function getAttributeAny(
  element: XmlElement | null | undefined,
  names: string[],
): string | null {
  if (!element || !element.attributes) {
    return null;
  }

  const attrs = element.attributes as Record<string, string>;

  for (const name of names) {
    if (name in attrs) {
      return attrs[name] ?? null;
    }
  }

  return null;
}

/**
 * Get all attributes from an element
 *
 * @param element - Element to get attributes from
 * @returns Record of attribute name -> value
 */
export function getAttributes(element: XmlElement | null | undefined): Record<string, string> {
  if (!element || !element.attributes) {
    return {};
  }
  return element.attributes as Record<string, string>;
}

/**
 * Get the text content of an element (concatenates all text nodes)
 *
 * @param element - Element to get text from
 * @returns Text content or empty string
 */
export function getTextContent(element: XmlElement | null | undefined): string {
  if (!element) {
    return "";
  }

  // Check for direct text property
  if ("text" in element && typeof element.text === "string") {
    return element.text;
  }

  // Check elements array for text nodes
  if (!element.elements) {
    return "";
  }

  let text = "";
  for (const child of element.elements) {
    if (child.type === "text" && "text" in child) {
      text += child.text ?? "";
    } else if (child.type === "element") {
      // Recurse into child elements
      text += getTextContent(child);
    }
  }

  return text;
}

/**
 * Check if an element has a specific attribute with value "true" or "1"
 *
 * @param element - Element to check
 * @param namespace - Attribute namespace
 * @param name - Attribute name
 * @returns true if attribute exists and is truthy
 */
export function hasFlag(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string,
): boolean {
  const value = getAttribute(element, namespace, name);

  // In OOXML, presence of element often means true, absence means false
  // If value is null, check if the element itself exists
  if (value === null) {
    return false;
  }

  // Explicitly false
  if (value === "0" || value === "false" || value === "off") {
    return false;
  }

  // Any other value (including "1", "true", "on", or empty string) means true
  return true;
}

/**
 * Check if a child element exists (used for boolean flags in OOXML)
 *
 * @param parent - Parent element
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns true if child element exists
 */
export function hasChild(
  parent: XmlElement | null | undefined,
  namespace: string,
  localName: string,
): boolean {
  return findChild(parent, namespace, localName) !== null;
}

/**
 * Parse an OOXML color value
 *
 * @param element - Color element (e.g., w:color)
 * @returns Object with val, themeColor, themeTint, themeShade
 */
export function parseColorElement(element: XmlElement | null | undefined): {
  val?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
} | null {
  if (!element) {
    return null;
  }

  const val = getAttribute(element, "w", "val");
  const themeColor = getAttribute(element, "w", "themeColor");
  const themeTint = getAttribute(element, "w", "themeTint");
  const themeShade = getAttribute(element, "w", "themeShade");
  return {
    ...(val != null ? { val } : {}),
    ...(themeColor != null ? { themeColor } : {}),
    ...(themeTint != null ? { themeTint } : {}),
    ...(themeShade != null ? { themeShade } : {}),
  };
}

/**
 * Parse a numeric value from an attribute, with optional scale
 *
 * @param element - Element containing the attribute
 * @param namespace - Attribute namespace
 * @param name - Attribute name
 * @param scale - Optional scale factor (e.g., 20 for twips to points)
 * @returns Parsed number or undefined
 */
export function parseNumericAttribute(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string,
  scale: number = 1,
): number | undefined {
  const value = getAttribute(element, namespace, name);
  if (value === null) {
    return undefined;
  }

  const num = Number.parseInt(value, 10);
  if (Number.isNaN(num)) {
    return undefined;
  }

  return num * scale;
}

/**
 * Parse a zero-based numbering level reference.
 *
 * OOXML numbering levels cannot be negative. Some producers emit negative
 * sentinel values alongside `w:numId="0"` to disable numbering; treat those
 * as an omitted level instead of leaking an invalid value into the model.
 */
export function parseNumberingLevelAttribute(
  element: XmlElement | null | undefined,
): number | undefined {
  const level = parseNumericAttribute(element, "w", "val");
  return level !== undefined && level >= 0 ? level : undefined;
}

/**
 * Parse `w:w` on a table width/height element. For `w:type="pct"`, producers
 * sometimes emit human-readable percentages (`100%`) instead of 50ths-of-percent
 * (`5000`); normalize those to the ECMA-376 unit the layout engine expects.
 */
export function parseTableMeasurementValue(
  element: XmlElement | null | undefined,
  widthType: string,
): number | undefined {
  const raw = getAttribute(element, "w", "w");
  if (raw === null) {
    return undefined;
  }

  const trimmed = raw.trim();

  if (widthType === "pct" && trimmed.endsWith("%")) {
    const pct = Number.parseFloat(trimmed.slice(0, -1));
    if (!Number.isNaN(pct)) {
      return Math.round(pct * 50);
    }
  }

  const num = Number.parseInt(trimmed, 10);
  return Number.isNaN(num) ? undefined : num;
}

/**
 * Parse a boolean value from an attribute or element presence
 *
 * OOXML boolean conventions:
 * - Element presence with no val attribute = true
 * - w:val="true" or w:val="1" = true
 * - w:val="false" or w:val="0" = false
 *
 * @param element - Element to check
 * @param namespace - Namespace for val attribute
 * @returns boolean value
 */
export function parseBooleanElement(
  element: XmlElement | null | undefined,
  namespace: string = "w",
): boolean {
  if (!element) {
    return false;
  }

  // OOXML binds prefixes to namespace URIs at the doc root; the source is
  // free to pick any prefix (`w14:checked` vs. `ns0:checked`) so long as
  // it resolves to the right URI. fast-xml-parser keeps prefixes literal,
  // so we have to be tolerant of the prefix actually written rather than
  // assuming the canonical one. Without this, a `<ns0:checked ns0:val="0"/>`
  // misses the val attribute, falls through to the bare-presence branch,
  // and an unchecked box renders as checked (codex P2, PR #587).
  let val: string | null = null;
  const elementName = element.name ?? "";
  const colonIdx = elementName.indexOf(":");
  const elementPrefix = colonIdx > 0 ? elementName.slice(0, colonIdx) : null;
  if (elementPrefix && elementPrefix !== namespace) {
    val = getAttribute(element, elementPrefix, "val");
  }
  if (val === null) {
    val = getAttribute(element, namespace, "val");
  }
  // A canonical element can carry the val attribute under a different
  // inherited prefix (`<w:b x:val="0"/>` when `x` is bound to the WP
  // URI). The wrapper-prefix and canonical-prefix lookups above miss
  // that; scan by local-name suffix so the OnOff parse is honest.
  if (val === null && element.attributes) {
    for (const [key, value] of Object.entries(element.attributes as Record<string, string>)) {
      if (key === "val" || key.endsWith(":val")) {
        val = value;
        break;
      }
    }
  }

  // No val attribute = true (element presence implies true)
  if (val === null) {
    return true;
  }

  // Explicit false values
  if (val === "0" || val === "false" || val === "off") {
    return false;
  }

  return true;
}

/**
 * Deep find - search recursively for an element
 *
 * @param root - Root element to search from
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns First matching element found or null
 */
export function findDeep(
  root: XmlElement | null | undefined,
  namespace: string,
  localName: string,
): XmlElement | null {
  if (!root) {
    return null;
  }

  // Check if this element matches
  if (matchesName(root, namespace, localName)) {
    return root;
  }

  // Search children
  if (root.elements) {
    for (const child of root.elements) {
      if (child.type !== "element") {
        continue;
      }

      const found = findDeep(child, namespace, localName);
      if (found) {
        return found;
      }
    }
  }

  return null;
}

/**
 * Find all elements matching name, searching recursively
 *
 * @param root - Root element to search from
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns Array of all matching elements
 */
export function findAllDeep(
  root: XmlElement | null | undefined,
  namespace: string,
  localName: string,
): XmlElement[] {
  const results: XmlElement[] = [];

  function search(element: XmlElement | null | undefined): void {
    if (!element) {
      return;
    }

    if (matchesName(element, namespace, localName)) {
      results.push(element);
    }

    if (element.elements) {
      for (const child of element.elements) {
        if (child.type === "element") {
          search(child);
        }
      }
    }
  }

  search(root);
  return results;
}

/**
 * Sanity cap on distinct `xmlns:*` declarations collected per element. Real
 * documents declare a handful; a hostile element with thousands of unique
 * prefixes would otherwise get replayed onto every captured `w:pict` rawXml
 * subtree that inherits from it.
 */
const MAX_XMLNS_DECLARATIONS_PER_ELEMENT = 64;

/**
 * Collect every `xmlns` / `xmlns:*` declaration from an element's attributes.
 *
 * The serializer's hard-coded root namespaces only cover canonical prefixes
 * (`a`, `w`, `wp`, `v`, `o`, …); a DOCX that binds a namespace to a
 * non-canonical prefix would replay an unbound prefix when a captured subtree
 * is emitted on its own. Carrying the source declarations forward keeps the
 * raw replay self-contained regardless of the producer's prefix choice.
 */
export function collectXmlnsDeclarations(element: XmlElement): Record<string, string> {
  const out: Record<string, string> = {};
  const attrs = element.attributes;
  if (!attrs) {
    return out;
  }
  let declarationCount = 0;
  for (const key in attrs) {
    if (declarationCount >= MAX_XMLNS_DECLARATIONS_PER_ELEMENT) {
      break;
    }
    const value = attrs[key];
    if ((key === "xmlns" || key.startsWith("xmlns:")) && value !== undefined) {
      out[key] = String(value);
      declarationCount += 1;
    }
  }
  return out;
}

/**
 * Merge an element's own `xmlns` / `xmlns:*` declarations onto an inherited
 * in-scope set, returning the accumulated set. Threaded down the ancestor
 * chain (root -> paragraph -> run -> pict), the element's own declarations
 * override an inherited binding for the same prefix, matching XML scoping. The
 * inherited object is not mutated; the same reference is returned unchanged when
 * the element declares nothing.
 */
export function mergeXmlnsDeclarations(
  inherited: Record<string, string>,
  element: XmlElement,
): Record<string, string> {
  const own = collectXmlnsDeclarations(element);
  for (const _key in own) {
    return { ...inherited, ...own };
  }
  return inherited;
}

const QNAME_VALUE_ATTRIBUTES = new Set([
  "Ignorable",
  "PreserveAttributes",
  "PreserveElements",
  "ProcessContent",
  "Requires",
]);

const addPrefix = (name: string, prefixes: Set<string>): void => {
  const separatorIndex = name.indexOf(":");
  if (separatorIndex > 0) {
    prefixes.add(name.slice(0, separatorIndex));
  }
};

/**
 * Collect namespace prefixes whose bindings a detached raw subtree needs.
 *
 * Element and attribute names use prefixes directly. Markup-compatibility and
 * QName-valued attributes can also refer to prefixes in their values, so keep
 * those bindings even when no descendant name uses them.
 */
const collectUsedNamespacePrefixes = (
  element: XmlElement,
  prefixes = new Set<string>(),
): Set<string> => {
  if (element.name) {
    if (element.name.includes(":")) {
      addPrefix(element.name, prefixes);
    } else {
      prefixes.add("");
    }
  }

  if (element.attributes) {
    for (const [name, value] of Object.entries(element.attributes)) {
      if (name === "xmlns" || name.startsWith("xmlns:")) {
        continue;
      }
      addPrefix(name, prefixes);
      const localName = name.slice(name.indexOf(":") + 1);
      if (name !== "xsi:type" && !QNAME_VALUE_ATTRIBUTES.has(localName)) {
        continue;
      }
      for (const token of String(value).split(/\s+/u)) {
        const separatorIndex = token.indexOf(":");
        prefixes.add(separatorIndex > 0 ? token.slice(0, separatorIndex) : token);
      }
    }
  }

  if (element.elements) {
    for (const child of element.elements) {
      if (child.type === "element") {
        collectUsedNamespacePrefixes(child, prefixes);
      }
    }
  }
  return prefixes;
};

const isCanonicalNamespacePrefix = (prefix: string): prefix is keyof typeof NAMESPACES =>
  Object.hasOwn(NAMESPACES, prefix);

/**
 * Return a shallow clone of `element` carrying the inherited namespace
 * declarations needed by its raw subtree.
 *
 * Existing declarations retain their authored order. Only missing, referenced
 * bindings are appended, which makes repeated detach/replay cycles idempotent
 * even when the surrounding document declares a different namespace superset.
 * If a producer omitted a binding for a standard OOXML prefix, use its
 * canonical binding so the detached subtree is valid on its first replay.
 */
export function cloneWithXmlnsDeclarations(
  element: XmlElement,
  xmlnsDecls: Record<string, string>,
): XmlElement {
  const prefixes = collectUsedNamespacePrefixes(element);
  const additions: Record<string, string> = {};
  const attributes = element.attributes ?? {};
  for (const [name, value] of Object.entries(xmlnsDecls)) {
    const prefix = name === "xmlns" ? "" : name.slice("xmlns:".length);
    if (prefixes.has(prefix) && attributes[name] === undefined) {
      additions[name] = value;
    }
  }
  for (const prefix of prefixes) {
    const name = `xmlns:${prefix}`;
    if (
      !prefix ||
      attributes[name] !== undefined ||
      additions[name] !== undefined ||
      !isCanonicalNamespacePrefix(prefix)
    ) {
      continue;
    }
    additions[name] = NAMESPACES[prefix];
  }
  for (const _key in additions) {
    return {
      ...element,
      attributes: { ...attributes, ...additions },
    };
  }
  return element;
}
