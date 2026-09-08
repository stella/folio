/**
 * The one way to turn parsed markup into a string a rebuilt part replays.
 *
 * folio parses Strict and Transitional packages and rebuilds every package as
 * Transitional. Modeled properties are re-serialized, so they come out
 * Transitional whatever the source was; content the parser keeps verbatim —
 * raw property XML, drawing and shape bodies, text-box markup, unmodeled
 * extensions — does not, and a Strict fragment copied under a Transitional root
 * changes meaning: the namespaces it declares stay Strict while the root is
 * not, and its values are spelled the way Strict spells them (`155.85pt`,
 * `20%`) where the Transitional attribute is a number.
 *
 * A fragment is therefore converted here, at capture, where its own namespace
 * scope still says which conformance class produced it. Nothing downstream has
 * to know: {@link captureVerbatimXml} is the only producer of replayed markup,
 * and a fragment that carries no Strict namespace passes through byte for byte.
 */

import { TaggedError } from "better-result";

import { assertXmlResourceLimits } from "./xmlResourceLimits";
import { getDocxXmlSafetyIssue } from "./xmlSafety";
import {
  type PercentUnit,
  type SlotEncoding,
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI,
} from "./strictValueEncodings.gen";
import {
  isStrictNamespaceUri,
  STRICT_URI_PREFIX,
  toTransitionalNamespaceUri,
  transitionalSlotEncoding,
} from "./transitionalSpelling";
import { roundHalfAwayFromZero, universalMeasureAs } from "./universalMeasure";
import {
  cloneElement,
  elementToXml,
  getChildElements,
  getLocalName,
  getNamespacePrefix,
  NAMESPACES,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  type XmlElement,
  type XmlNamespaceScope,
} from "./xmlParser";

/** A Strict fragment uses a namespace that has no Transitional counterpart. */
export class UntranslatableStrictNamespaceError extends TaggedError(
  "UntranslatableStrictNamespaceError",
)<{
  message: string;
  uri: string;
}> {}

const TRANSITIONAL_URIS: ReadonlySet<string> = new Set(
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI.values(),
);
const XML_NAMESPACE_URI = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE_URI = "http://www.w3.org/2000/xmlns/";
const XSI_NAMESPACE_URI = "http://www.w3.org/2001/XMLSchema-instance";
const NCNAME = /^[\p{L}_][\p{L}\p{N}._\-\u00b7\p{M}]*$/u;
const MC_QNAME_LIST_ATTRIBUTES = new Set([
  "PreserveAttributes",
  "PreserveElements",
  "ProcessContent",
]);

type NamespaceValueKind = "prefix-list" | "qname" | "qname-list";

/**
 * Drawing extensions the repository has no schema for.
 *
 * `wp14:pctWidth` and its siblings are Microsoft vocabularies, but they are
 * DrawingML by construction and DrawingML has one integer encoding for a
 * percentage, so a Strict producer's `20%` there converts like `a:alpha`'s.
 * Lengths are left alone: only an attribute's own type says what unit its
 * number counts, and guessing one would corrupt the value silently.
 */
const DRAWING_EXTENSION_MARKERS: readonly string[] = [
  "/wordprocessingCanvas",
  "/wordprocessingDrawing",
  "/wordprocessingGroup",
  "/wordprocessingInk",
  "/wordprocessingShape",
];

const isUnschemadDrawingNamespace = (uri: string): boolean =>
  !TRANSITIONAL_URIS.has(uri) && DRAWING_EXTENSION_MARKERS.some((marker) => uri.includes(marker));

/** `-?12.5%`: the one shape ECMA-376 gives a percentage that carries its unit. */
const PERCENTAGE = /^(-?[0-9]+(?:\.[0-9]+)?)%$/u;

const NUMBERS_PER_PERCENT: Readonly<Record<PercentUnit, number>> = {
  fiftiethPercent: 50,
  thousandthPercent: 1000,
  wholePercent: 1,
};

/** The Transitional spelling of one Strict-produced value, or the value unchanged. */
const transitionalValue = (
  value: string,
  namespaceUri: string,
  encoding: SlotEncoding | undefined,
): string => {
  const measureUnit = encoding?.measure;
  if (measureUnit !== undefined) {
    const measure = universalMeasureAs(value, measureUnit);
    if (measure !== undefined) {
      return String(measure);
    }
  }

  const percentage = PERCENTAGE.exec(value);
  if (percentage === null) {
    return value;
  }
  const percentUnit =
    encoding?.percent ??
    (isUnschemadDrawingNamespace(namespaceUri) ? "thousandthPercent" : undefined);
  if (percentUnit === undefined) {
    return value;
  }
  // SAFETY: the capture group is present whenever the pattern matched.
  return String(roundHalfAwayFromZero(Number(percentage[1]!) * NUMBERS_PER_PERCENT[percentUnit]));
};

/**
 * Which conformance class produced a subtree's markup.
 *
 * Inherited rather than read per element: an extension such as
 * `wp14:sizeRelH` has one namespace in both classes, so only its OOXML
 * ancestry says how its producer spelled the values inside it.
 */
type FragmentOrigin = "strict" | "transitional";

const originOf = (element: XmlElement, inherited: FragmentOrigin): FragmentOrigin => {
  const uri = element.namespaceUri;
  if (uri === undefined) {
    return inherited;
  }
  if (isStrictNamespaceUri(uri)) {
    return "strict";
  }
  if (uri.startsWith(STRICT_URI_PREFIX)) {
    throw new UntranslatableStrictNamespaceError({
      message: `Strict namespace ${uri} has no Transitional counterpart; folio will not replay it under a Transitional root.`,
      uri,
    });
  }
  return TRANSITIONAL_URIS.has(uri) ? "transitional" : inherited;
};

const convertAttributes = (
  element: XmlElement,
  namespaceUri: string,
  localName: string,
  convertValues: boolean,
): Record<string, string | number | undefined> | undefined => {
  const attributes = element.attributes ?? {};
  let converted: Record<string, string | number | undefined> | undefined;
  for (const [name, raw] of Object.entries(attributes)) {
    if (typeof raw !== "string") {
      continue;
    }
    let next = raw;
    if (name === "xmlns" || name.startsWith("xmlns:")) {
      next = toTransitionalNamespaceUri(raw);
    } else if (convertValues) {
      next = transitionalValue(
        raw,
        namespaceUri,
        transitionalSlotEncoding(element.namespaceUri, localName, getLocalName(name)),
      );
    }
    if (next === raw) {
      continue;
    }
    converted ??= { ...attributes };
    converted[name] = next;
  }
  return converted;
};

const childWithValue = (child: XmlElement, text: string): XmlElement =>
  text === child.text ? child : { ...child, text };

const toTransitional = (element: XmlElement, inherited: FragmentOrigin): XmlElement => {
  if (element.type === "text") {
    return element;
  }

  const origin = originOf(element, inherited);
  const namespaceUri =
    element.namespaceUri === undefined ? "" : toTransitionalNamespaceUri(element.namespaceUri);
  const localName = getLocalName(element.name);
  const textEncoding = transitionalSlotEncoding(element.namespaceUri, localName);

  const attributes = element.attributes
    ? convertAttributes(element, namespaceUri, localName, origin === "strict")
    : undefined;

  let elements: XmlElement[] | undefined;
  for (const [index, child] of (element.elements ?? []).entries()) {
    const next =
      origin === "strict" && child.type === "text" && typeof child.text === "string"
        ? childWithValue(child, transitionalValue(child.text, namespaceUri, textEncoding))
        : toTransitional(child, origin);
    if (next === child) {
      continue;
    }
    elements ??= [...(element.elements ?? [])];
    elements[index] = next;
  }

  if (attributes === undefined && elements === undefined) {
    return element;
  }
  return {
    ...element,
    ...(attributes === undefined ? {} : { attributes }),
    ...(elements === undefined ? {} : { elements }),
  };
};

/** True when any binding in scope, including inherited ones, is a Strict namespace. */
const scopeDeclaresStrict = (scope: XmlNamespaceScope | undefined): boolean => {
  for (let current = scope; current; current = current.parent) {
    for (const uri of current.bindings.values()) {
      if (uri.startsWith(STRICT_URI_PREFIX)) {
        return true;
      }
    }
  }
  return false;
};

const resolveNamespaceBinding = (
  scope: XmlNamespaceScope | undefined,
  prefix: string,
): string | undefined => {
  if (prefix === "xml") {
    return XML_NAMESPACE_URI;
  }
  for (let current = scope; current; current = current.parent) {
    const value = current.bindings.get(prefix);
    if (value !== undefined) {
      return value;
    }
  }
  return undefined;
};

const namespaceValueKind = (
  element: XmlElement,
  attributeName: string,
): NamespaceValueKind | null => {
  const prefix = getNamespacePrefix(attributeName);
  const namespaceUri =
    prefix === null
      ? ""
      : toTransitionalNamespaceUri(resolveNamespaceBinding(element.namespaceScope, prefix) ?? "");
  const localName = getLocalName(attributeName);
  if (namespaceUri === XSI_NAMESPACE_URI && localName === "type") {
    return "qname";
  }
  if (namespaceUri === NAMESPACES.mc) {
    if (localName === "Ignorable" || localName === "MustUnderstand") {
      return "prefix-list";
    }
    if (MC_QNAME_LIST_ATTRIBUTES.has(localName)) {
      return "qname-list";
    }
  }
  if (
    prefix === null &&
    localName === "Requires" &&
    getLocalName(element.name ?? "") === "Choice" &&
    toTransitionalNamespaceUri(element.namespaceUri ?? "") === NAMESPACES.mc
  ) {
    return "prefix-list";
  }
  return null;
};

const referencedNamespacePrefixes = (
  element: XmlElement,
  attributeName: string,
  value: string,
): readonly string[] | null => {
  const kind = namespaceValueKind(element, attributeName);
  if (kind === null) {
    return [];
  }
  const tokens = value.trim() === "" ? [] : value.trim().split(/\s+/u);
  if (kind === "qname" && tokens.length !== 1) {
    return null;
  }
  const prefixes: string[] = [];
  for (const token of tokens) {
    if (kind === "prefix-list") {
      if (
        !NCNAME.test(token) ||
        resolveNamespaceBinding(element.namespaceScope, token) === undefined
      ) {
        return null;
      }
      prefixes.push(token);
      continue;
    }
    const parts = token.split(":");
    if (parts.length > 2 || parts.some((part) => !NCNAME.test(part))) {
      return null;
    }
    const prefix = parts.length === 2 ? parts[0]! : "";
    if (prefix !== "" && resolveNamespaceBinding(element.namespaceScope, prefix) === undefined) {
      return null;
    }
    prefixes.push(prefix);
  }
  return prefixes;
};

/**
 * Serialize parsed markup for replay inside a rebuilt Transitional part.
 *
 * Every verbatim capture goes through here — a lint rule keeps
 * {@link elementToXml} out of reach of the parsers — so a replay path added
 * later cannot skip the conversion. Markup that binds no Strict namespace,
 * anywhere in scope or in the fragment itself, is serialized once and returned
 * unchanged.
 */
const materializeInheritedNamespaceBindings = (element: XmlElement): XmlElement => {
  const additions: Record<string, string> = {};
  const rootAttributes = element.attributes ?? {};

  const addInheritedBinding = (node: XmlElement, prefix: string): void => {
    if (prefix === "xml" || prefix === "xmlns") {
      return;
    }
    const attributeName = prefix === "" ? "xmlns" : `xmlns:${prefix}`;
    if (rootAttributes[attributeName] !== undefined || additions[attributeName] !== undefined) {
      return;
    }
    const namespaceUri = resolveNamespaceBinding(node.namespaceScope, prefix);
    if (namespaceUri === undefined) {
      return;
    }
    const canonicalNamespace = OOXML_NAMESPACE_SCOPE.bindings.get(prefix);
    if (
      canonicalNamespace !== undefined &&
      toTransitionalNamespaceUri(namespaceUri) === canonicalNamespace
    ) {
      return;
    }
    additions[attributeName] = namespaceUri;
  };

  const visit = (node: XmlElement, inheritedDeclarations: ReadonlySet<string>): void => {
    const declarations = new Set(inheritedDeclarations);
    for (const name of Object.keys(node.attributes ?? {})) {
      if (name === "xmlns") {
        declarations.add("");
      } else if (name.startsWith("xmlns:")) {
        declarations.add(name.slice("xmlns:".length));
      }
    }

    const elementPrefix = node.name ? getNamespacePrefix(node.name) : null;
    if (elementPrefix !== null && !declarations.has(elementPrefix)) {
      addInheritedBinding(node, elementPrefix);
    } else if (elementPrefix === null && !declarations.has("")) {
      addInheritedBinding(node, "");
    }

    for (const [name, value] of Object.entries(node.attributes ?? {})) {
      if (name === "xmlns" || name.startsWith("xmlns:")) {
        continue;
      }
      const attributePrefix = getNamespacePrefix(name);
      if (attributePrefix !== null && !declarations.has(attributePrefix)) {
        addInheritedBinding(node, attributePrefix);
      }
      const prefixes = referencedNamespacePrefixes(node, name, String(value));
      if (prefixes === null) {
        continue;
      }
      for (const prefix of prefixes) {
        if (!declarations.has(prefix)) {
          addInheritedBinding(node, prefix);
        }
      }
    }

    for (const child of node.elements ?? []) {
      if (child.type === "element") {
        visit(child, declarations);
      }
    }
  };

  visit(element, new Set());
  return Object.keys(additions).length === 0
    ? element
    : cloneElement(element, { attributes: { ...rootAttributes, ...additions } });
};

export const captureVerbatimXml = (element: XmlElement): string => {
  const replayableElement = materializeInheritedNamespaceBindings(element);
  const xml = elementToXml(replayableElement);
  if (!xml.includes(STRICT_URI_PREFIX) && !scopeDeclaresStrict(replayableElement.namespaceScope)) {
    return xml;
  }
  return elementToXml(toTransitional(replayableElement, "transitional"));
};

type SanitizeCapturedXmlOptions = {
  allowedLocalNames: ReadonlySet<string>;
  allowedNamespaceUris: ReadonlySet<string>;
  inheritedNamespaceScope?: XmlNamespaceScope;
  requiredNamespaceBindings?: ReadonlyMap<string, string>;
  transform?: (element: XmlElement) => XmlElement;
  validate?: (element: XmlElement) => boolean;
};

const hasRequiredNamespaceBindings = (
  element: XmlElement,
  required: ReadonlyMap<string, string> | undefined,
): boolean => {
  if (!required) {
    return true;
  }
  for (const [prefix, namespaceUri] of required) {
    const resolved = resolveNamespaceBinding(element.namespaceScope, prefix);
    if (resolved === undefined || toTransitionalNamespaceUri(resolved) !== namespaceUri) {
      return false;
    }
  }
  return true;
};

const validNamespaceDeclaration = (name: string, value: string): boolean => {
  if (name === "xmlns") {
    return value !== XML_NAMESPACE_URI && value !== XMLNS_NAMESPACE_URI;
  }
  const prefix = name.slice("xmlns:".length);
  const normalizedPrefix = prefix.toLowerCase();
  if (normalizedPrefix.startsWith("xml") && prefix !== "xml") {
    return false;
  }
  if (prefix === "xml") {
    return value === XML_NAMESPACE_URI;
  }
  return value !== "" && value !== XML_NAMESPACE_URI && value !== XMLNS_NAMESPACE_URI;
};

const hasBoundNamespaces = (root: XmlElement): boolean => {
  const pending = [root];
  while (pending.length > 0) {
    const element = pending.pop();
    if (!element || element.type !== "element" || !element.name) {
      return false;
    }
    const elementPrefix = getNamespacePrefix(element.name);
    if (
      elementPrefix === "xmlns" ||
      (elementPrefix !== null &&
        elementPrefix !== "xml" &&
        resolveNamespaceBinding(element.namespaceScope, elementPrefix) === undefined)
    ) {
      return false;
    }

    const expandedAttributeNames = new Set<string>();
    for (const [name, rawValue] of Object.entries(element.attributes ?? {})) {
      const value = String(rawValue ?? "");
      if (name === "xmlns" || name.startsWith("xmlns:")) {
        if (!validNamespaceDeclaration(name, value)) {
          return false;
        }
        continue;
      }
      const prefix = getNamespacePrefix(name);
      if (
        prefix === "xmlns" ||
        (prefix !== null &&
          prefix !== "xml" &&
          resolveNamespaceBinding(element.namespaceScope, prefix) === undefined)
      ) {
        return false;
      }
      const attributeNamespace =
        prefix === null ? "" : resolveNamespaceBinding(element.namespaceScope, prefix);
      // Strict and Transitional bindings become the same namespace on replay.
      // Detect the resulting duplicate before conversion so two lexical names
      // cannot collapse into one expanded attribute name.
      const replayNamespace = toTransitionalNamespaceUri(attributeNamespace ?? "");
      const expandedName = `${replayNamespace}\u0000${getLocalName(name)}`;
      if (expandedAttributeNames.has(expandedName)) {
        return false;
      }
      expandedAttributeNames.add(expandedName);
      if (referencedNamespacePrefixes(element, name, value) === null) {
        return false;
      }
    }
    for (const child of getChildElements(element)) {
      pending.push(child);
    }
  }
  return true;
};

const parsedReplayEnvelope = (
  xml: string,
  inheritedNamespaceScope: XmlNamespaceScope,
): XmlElement | null => {
  let wrapper: XmlElement | undefined;
  try {
    wrapper = parseXml(
      `<folio-replay-envelope>${xml}</folio-replay-envelope>`,
      inheritedNamespaceScope,
    ).elements?.at(0);
  } catch {
    return null;
  }
  if (!wrapper || wrapper.type !== "element") {
    return null;
  }
  let root: XmlElement | null = null;
  for (const child of wrapper.elements ?? []) {
    if (child.type === "text") {
      if (String(child.text ?? "").trim() !== "") {
        return null;
      }
      continue;
    }
    if (child.type !== "element" || root !== null) {
      return null;
    }
    root = child;
  }
  return root;
};

const withinXmlResourceLimits = (xml: string): boolean => {
  try {
    assertXmlResourceLimits(xml);
    return true;
  } catch {
    return false;
  }
};

const EMPTY_NAMESPACE_SCOPE: XmlNamespaceScope = { bindings: new Map() };

/** Validate a complete XML part immediately before package output. */
export const isSafeCapturedXmlDocument = (xml: string): boolean => {
  const trimmed = xml.trim();
  if (trimmed === "" || !withinXmlResourceLimits(xml) || getDocxXmlSafetyIssue(xml) !== null) {
    return false;
  }
  const root = parsedReplayEnvelope(trimmed, EMPTY_NAMESPACE_SCOPE);
  return root !== null && hasBoundNamespaces(root);
};

const captureAllowedElement = (
  element: XmlElement,
  {
    allowedLocalNames,
    allowedNamespaceUris,
    inheritedNamespaceScope,
    requiredNamespaceBindings,
    transform,
    validate,
  }: SanitizeCapturedXmlOptions,
): string | null => {
  let passesDomainValidation = true;
  try {
    passesDomainValidation = validate?.(element) ?? true;
  } catch {
    return null;
  }
  if (
    element.type !== "element" ||
    !element.name ||
    !allowedLocalNames.has(getLocalName(element.name)) ||
    !allowedNamespaceUris.has(toTransitionalNamespaceUri(element.namespaceUri ?? "")) ||
    !hasBoundNamespaces(element) ||
    !hasRequiredNamespaceBindings(element, requiredNamespaceBindings) ||
    !passesDomainValidation
  ) {
    return null;
  }
  try {
    const replayElement = transform?.(element) ?? element;
    const transformedPassesDomainValidation =
      transform === undefined ? true : (validate?.(replayElement) ?? true);
    if (
      !transformedPassesDomainValidation ||
      !hasBoundNamespaces(replayElement) ||
      !hasRequiredNamespaceBindings(replayElement, requiredNamespaceBindings)
    ) {
      return null;
    }
    const captured = captureVerbatimXml(replayElement);
    const capturedRoot = parsedReplayEnvelope(
      captured,
      inheritedNamespaceScope ?? OOXML_NAMESPACE_SCOPE,
    );
    if (
      capturedRoot === null ||
      !hasBoundNamespaces(capturedRoot) ||
      !hasRequiredNamespaceBindings(capturedRoot, requiredNamespaceBindings) ||
      !(validate?.(capturedRoot) ?? true)
    ) {
      return null;
    }
    return captured;
  } catch {
    return null;
  }
};

/**
 * Re-parse and serialize one captured element before replaying it.
 *
 * The parser deliberately ignores declarations, processing instructions,
 * comments, and document types. Returning the caller's source string after a
 * successful parse would therefore replay tokens that were never validated.
 * This function returns only the parsed root serialized through the capture
 * owner, and rejects siblings, unexpected roots, and namespace shadowing.
 */
export const sanitizeCapturedXmlElement = (
  xml: string | undefined,
  options: SanitizeCapturedXmlOptions,
): string | null => {
  if (xml === undefined) {
    return null;
  }
  const trimmed = xml.trim();
  if (!trimmed) {
    return null;
  }
  if (!withinXmlResourceLimits(xml) || getDocxXmlSafetyIssue(xml) !== null) {
    return null;
  }

  const root = parsedReplayEnvelope(
    trimmed,
    options.inheritedNamespaceScope ?? OOXML_NAMESPACE_SCOPE,
  );
  return root ? captureAllowedElement(root, options) : null;
};
