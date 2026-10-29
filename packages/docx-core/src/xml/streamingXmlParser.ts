/**
 * folio's one XML reader: a single forward pass that builds the shared
 * element tree, or scans start tags without building one.
 *
 * Every reader in the workspace goes through it, and each states how much it
 * accepts:
 *
 * - `ooxml`: the ordinary contract the DOCX readers parse with. The five
 *   built-in entities and character references are decoded; a DOCTYPE, any
 *   other entity, mismatched or unclosed markup, text outside the root and
 *   nesting past {@link FOLIO_XML_RESOURCE_LIMITS} are refused.
 * - `well-formed`: XML 1.0 well-formedness on top of that, for validators. One
 *   root element, valid names, no repeated attribute, no `<` in an attribute
 *   value, references only to characters XML admits, comments and processing
 *   instructions in their lexical form, and no characters XML excludes.
 * - `tolerant`: `ooxml`, except that text outside the root is dropped and a
 *   reference the `ooxml` reader refuses is read the way earlier folio
 *   releases read it (common HTML names decode, an unknown name stays
 *   literal), so a part those releases opened still opens with the same text.
 *
 * `references: "raw"` keeps each reference as written instead of decoding it,
 * for readers that compare or re-emit source spellings.
 */
import { TaggedError, type TaggedErrorClass } from "better-result";

import { isOoxmlNumericIdAttributeName } from "../serialize/numericIdNames";
import type { XmlElement, XmlNamespaceScope } from "./xmlElement";
import {
  attachXmlNamespaceContext,
  EMPTY_NAMESPACE_SCOPE,
} from "./xmlNamespaceContext";
import { FOLIO_XML_RESOURCE_LIMITS } from "./xmlResourceLimits";

export type { XmlElement, XmlNamespaceScope } from "./xmlElement";

type ParseXmlResult =
  { status: "parsed"; value: XmlElement } | { status: "unsupported" };

type ElementFrame = {
  element: XmlElement;
  name: string;
  start?: number;
};

type XmlSourceRange = { readonly start: number; readonly end: number };
const sourceRanges = new WeakMap<XmlElement, XmlSourceRange>();

/** Exact element source ranges, available only from the opt-in source-range reader. */
export const getXmlSourceRange = (
  element: XmlElement,
): XmlSourceRange | undefined => sourceRanges.get(element);

const BUILT_IN_ENTITIES = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"',
} as const;

/** How much markup a read accepts; see the module comment. */
export type XmlConformance = "ooxml" | "well-formed" | "tolerant";

/** Whether references are decoded or kept as written. */
export type XmlReferenceHandling = "decode" | "raw";

/**
 * What a reference to an entity no declaration defines does. `refuse` is
 * XML 1.0: without a DTD only the five built-ins exist. `accept` admits any
 * reference that is syntactically one (`&Name;`, `&#digits;`, `&#xhex;`) and
 * keeps it as written, for checks that must accept what earlier releases
 * accepted. Not used by the tolerant reader, which accepts every reference.
 */
export type XmlUndeclaredEntities = "refuse" | "accept";

const XmlReadErrorBase: TaggedErrorClass<"XmlReadError"> =
  TaggedError("XmlReadError");

/** An XML part the reader refused under the conformance its caller asked for. */
export class XmlReadError extends XmlReadErrorBase<{
  message: string;
  reason: string;
}> {}

type AttributeValueSpan = { start: number; end: number };
type OpenTagVisitor = (
  element: XmlElement,
  attributeValueSpans: ReadonlyMap<string, AttributeValueSpan>,
) => ReadonlyMap<string, string> | null;

type OpenTagScanVisitor = (...args: Parameters<OpenTagVisitor>) => null;

type XmlReplacement = AttributeValueSpan & { value: string };
type Rejected = { status: "unsupported"; reason: string };
type InternalParseXmlResult =
  | { status: "parsed"; value: XmlElement; replacements: XmlReplacement[] }
  | Rejected;

const reject = (reason: string): Rejected => ({
  status: "unsupported",
  reason,
});

type AttributeNameOptions = { attributeName: string; elementName: string };

type ElementScanOptions = {
  element: XmlElement;
  attributeValueSpans: ReadonlyMap<string, AttributeValueSpan>;
  nameEnd: number;
  parent: XmlElement | undefined;
};
type ElementScanVisitor = (options: ElementScanOptions) => void;

type StreamingXmlOptions = {
  xml: string;
  visitElement?: ElementScanVisitor;
  inheritedNamespaceScope?: XmlNamespaceScope;
  conformance?: XmlConformance;
  references?: XmlReferenceHandling;
  undeclaredEntities?: XmlUndeclaredEntities;
  instructions?: "skip" | "retain";
} & (
  | {
      mode: "tree";
      visitOpenTag?: OpenTagScanVisitor;
      sourceRanges?: "tracked";
    }
  | {
      mode: "attributes";
      visitOpenTag: OpenTagVisitor;
      retainAttribute?: (options: AttributeNameOptions) => boolean;
    }
  | {
      /** Only the verdict: no tree, no namespace resolution, no visitors. */
      mode: "verdict";
      visitOpenTag?: undefined;
    }
);

/** Reads a text run or attribute value; `null` refuses it. */
type ValueReader = (value: string) => string | null;

type ReadRules = {
  conformance: XmlConformance;
  readValue: ValueReader;
  /** Names already checked against the XML `Name` production in this read. */
  validNames: Set<string> | undefined;
};

type ReadPolicy = {
  conformance: XmlConformance;
  references: XmlReferenceHandling;
  undeclaredEntities: XmlUndeclaredEntities;
};

const valueReaderFor = ({
  conformance,
  references,
  undeclaredEntities,
}: ReadPolicy): ValueReader => {
  if (conformance === "tolerant") {
    return references === "decode"
      ? (value) => decodeTolerantReferences(normalizeLineEndings(value))
      : normalizeLineEndings;
  }
  let decodeReference = decodeOoxmlReference;
  if (undeclaredEntities === "accept") {
    decodeReference = decodeReferenceKeepingUndeclared;
  } else if (conformance === "well-formed") {
    decodeReference = decodeWellFormedReference;
  }
  return references === "decode"
    ? (value) => decodeReferences(normalizeLineEndings(value), decodeReference)
    : (value) => {
        const normalized = normalizeLineEndings(value);
        return decodeReferences(normalized, decodeReference) === null
          ? null
          : normalized;
      };
};

const readRulesFor = (policy: ReadPolicy): ReadRules => ({
  conformance: policy.conformance,
  readValue: valueReaderFor(policy),
  validNames: policy.conformance === "well-formed" ? new Set() : undefined,
});

// Built on first use: the reference decoders are declared further down.
let defaultRules: ReadRules | undefined;

const parseStreamingXmlInternal = (
  options: StreamingXmlOptions,
): InternalParseXmlResult => {
  const { xml, inheritedNamespaceScope = EMPTY_NAMESPACE_SCOPE } = options;
  const rules =
    options.conformance === undefined &&
    options.references === undefined &&
    options.undeclaredEntities === undefined
      ? (defaultRules ??= readRulesFor({
          conformance: "ooxml",
          references: "decode",
          undeclaredEntities: "refuse",
        }))
      : readRulesFor({
          conformance: options.conformance ?? "ooxml",
          references: options.references ?? "decode",
          undeclaredEntities: options.undeclaredEntities ?? "refuse",
        });
  const wellFormed = rules.conformance === "well-formed";
  const retainInstructions = options.instructions === "retain";
  if (wellFormed && !containsOnlyXmlCharacters(xml)) {
    return reject("a character XML does not allow");
  }
  const visitOpenTag = options.visitOpenTag;
  let spanMode: ParseOpenTagOptions["spanMode"] = "none";
  if (visitOpenTag !== undefined)
    spanMode = options.mode === "tree" ? "identities" : "all";
  if (options.visitElement !== undefined) spanMode = "all";
  const openTagOptions = {
    xml,
    start: 0,
    close: 0,
    spanMode,
    rules,
    retainAttribute:
      options.mode === "attributes" ? options.retainAttribute : undefined,
  } satisfies ParseOpenTagOptions;
  const root: XmlElement = { elements: [] };
  const stack: ElementFrame[] = [];
  const replacements: XmlReplacement[] = [];
  const documentStart = xml.charCodeAt(0) === 0xfeff ? 1 : 0;
  let rootElements = 0;
  let cursor = 0;
  let mergeAdjacentText = false;
  const acceptText = (text: string, merge: boolean): boolean => {
    if (wellFormed && text.includes("]]>")) return false;
    if (options.mode === "tree") return appendText(text, stack, merge, rules);
    const decoded = rules.readValue(text);
    return (
      decoded !== null &&
      (stack.length > 0 ||
        isXmlWhitespaceText(decoded) ||
        rules.conformance === "tolerant")
    );
  };

  while (cursor < xml.length) {
    const open = xml.indexOf("<", cursor);
    if (open === -1) {
      if (!acceptText(xml.slice(cursor), mergeAdjacentText)) {
        return reject("text the reader does not accept");
      }
      break;
    }

    if (
      open > cursor &&
      !acceptText(xml.slice(cursor, open), mergeAdjacentText)
    ) {
      return reject("text the reader does not accept");
    }
    if (open > cursor) {
      mergeAdjacentText = true;
    }

    if (xml.startsWith("<!--", open)) {
      const close = xml.indexOf("-->", open + 4);
      if (close === -1) {
        return reject("an unclosed comment");
      }
      if (wellFormed) {
        const body = xml.slice(open + 4, close);
        if (body.includes("--") || body.endsWith("-")) {
          return reject("a comment containing --");
        }
      }
      cursor = close + 3;
      continue;
    }

    if (xml.startsWith("<![CDATA[", open)) {
      const close = xml.indexOf("]]>", open + 9);
      if (close === -1) {
        return reject("an unclosed CDATA section");
      }
      // XML 1.0 admits CDATA only inside the root element, whatever it holds.
      if (wellFormed && stack.length === 0) {
        return reject("a CDATA section outside the root element");
      }
      const text = normalizeLineEndings(xml.slice(open + 9, close));
      if (
        !(options.mode === "tree"
          ? appendRawText(text, stack, rules)
          : stack.length > 0 ||
            isXmlWhitespaceText(text) ||
            rules.conformance === "tolerant")
      ) {
        return reject("text outside the root element");
      }
      cursor = close + 3;
      mergeAdjacentText = false;
      continue;
    }

    if (xml.startsWith("<?", open)) {
      const close = xml.indexOf("?>", open + 2);
      if (close === -1) {
        return reject("an unclosed processing instruction");
      }
      const instruction = readInstruction(xml, open, close);
      if (wellFormed) {
        const issue = instructionIssue({
          instruction,
          open,
          documentStart,
          rules,
        });
        if (issue !== null) return reject(issue);
      }
      if (retainInstructions && options.mode === "tree") {
        appendElement(stack.at(-1)?.element ?? root, {
          type: "instruction",
          name: instruction.target,
          instruction: instruction.body,
        });
      }
      cursor = close + 2;
      // Text ends at a processing instruction, though
      // not at a comment; the text on either side stays two nodes, as folio's
      // readers have always seen it.
      mergeAdjacentText = false;
      continue;
    }

    if (xml.startsWith("<!", open)) {
      return reject("a DOCTYPE or markup declaration");
    }

    const close = findTagClose(xml, open + 1);
    if (close === -1) {
      return reject("an unclosed tag");
    }

    if (xml.charCodeAt(open + 1) === 47) {
      if (wellFormed && isXmlWhitespace(xml.charCodeAt(open + 2))) {
        return reject("whitespace before an end-tag name");
      }
      const name = xml.slice(open + 2, close).trim();
      const frame = stack.pop();
      if (!frame || frame.name !== name) {
        return reject("an end tag that does not match its start tag");
      }
      if (frame.start !== undefined) {
        sourceRanges.set(frame.element, { start: frame.start, end: close + 1 });
      }
      cursor = close + 1;
      mergeAdjacentText = false;
      continue;
    }

    openTagOptions.start = open + 1;
    openTagOptions.close = close;
    const parsedTag = parseOpenTag(openTagOptions);
    if (parsedTag.status === "unsupported") {
      return parsedTag;
    }

    if (stack.length === 0) {
      rootElements += 1;
      if (wellFormed && rootElements > 1) {
        return reject("more than one root element");
      }
    }
    const parent = stack.at(-1)?.element;
    const scope = parent?.namespaceScope ?? inheritedNamespaceScope;
    if (options.mode === "verdict") {
      // Nothing reads the element; only its nesting is tracked below.
    } else if (
      options.mode === "attributes" &&
      options.retainAttribute !== undefined &&
      parsedTag.element.attributes === undefined
    ) {
      parsedTag.element.namespaceScope = scope;
    } else {
      attachXmlNamespaceContext(parsedTag.element, scope);
    }
    options.visitElement?.({
      element: parsedTag.element,
      attributeValueSpans: parsedTag.attributeValueSpans ?? new Map(),
      nameEnd: parsedTag.nameEnd,
      parent,
    });
    const shouldVisit =
      options.mode === "tree"
        ? parsedTag.attributeValueSpans !== undefined
        : options.mode === "attributes" &&
          (options.retainAttribute === undefined ||
            parsedTag.element.attributes !== undefined);
    const rewritten = shouldVisit
      ? visitOpenTag?.(
          parsedTag.element,
          parsedTag.attributeValueSpans ?? new Map(),
        )
      : null;
    if (rewritten) {
      for (const [attributeName, value] of rewritten) {
        const span = parsedTag.attributeValueSpans?.get(attributeName);
        if (!span) {
          return reject("a rewrite of an attribute the tag does not carry");
        }
        replacements.push({ ...span, value });
      }
    }
    // A self-closing element sits at the same depth an open one would, so
    // the limit applies to both, as the resource preflight counts them.
    if (stack.length >= FOLIO_XML_RESOURCE_LIMITS.maxDepth) {
      return reject(
        `nesting deeper than ${String(FOLIO_XML_RESOURCE_LIMITS.maxDepth)} elements`,
      );
    }
    if (options.mode === "tree")
      appendElement(parent ?? root, parsedTag.element);
    if (!parsedTag.selfClosing) {
      stack.push(
        options.mode === "tree" && options.sourceRanges === "tracked"
          ? { element: parsedTag.element, name: parsedTag.name, start: open }
          : { element: parsedTag.element, name: parsedTag.name },
      );
    } else if (options.mode === "tree" && options.sourceRanges === "tracked") {
      sourceRanges.set(parsedTag.element, { start: open, end: close + 1 });
    }
    cursor = close + 1;
    mergeAdjacentText = false;
  }

  if (stack.length > 0) {
    return reject("an unclosed element");
  }
  if (wellFormed && rootElements === 0) {
    return reject("no root element");
  }
  return { status: "parsed", value: root, replacements };
};

/** What a read returned: the tree, or why the reader refused the input. */
export type XmlReadResult =
  | { status: "parsed"; value: XmlElement }
  | { status: "rejected"; reason: string };

export type XmlReadOptions = {
  conformance: XmlConformance;
  /** Defaults to `"decode"`. */
  references?: XmlReferenceHandling;
  /** Defaults to `"refuse"`. */
  undeclaredEntities?: XmlUndeclaredEntities;
  /**
   * `"retain"` keeps processing instructions, the XML declaration included,
   * as `{ type: "instruction", name, instruction }` nodes; comments are always
   * dropped.
   */
  instructions?: "skip" | "retain";
  /** The bindings in scope around `xml`, for a fragment captured out of a part. */
  inheritedNamespaceScope?: XmlNamespaceScope;
};

/** Read `xml` into a tree under the stated conformance, in one pass. */
export const readXml = (
  xml: string,
  {
    conformance,
    references = "decode",
    undeclaredEntities = "refuse",
    instructions = "skip",
    inheritedNamespaceScope = EMPTY_NAMESPACE_SCOPE,
  }: XmlReadOptions,
): XmlReadResult => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "tree",
    conformance,
    references,
    undeclaredEntities,
    instructions,
    inheritedNamespaceScope,
  });
  return parsed.status === "parsed"
    ? { status: "parsed", value: parsed.value }
    : { status: "rejected", reason: parsed.reason };
};

/**
 * The reason `xml` is not well-formed XML, or `null` when it is.
 *
 * Stricter than XML 1.0 in one respect: an element or attribute named like a
 * built-in object property (`__proto__`, `constructor`, `toString`, ...) is
 * refused even though XML allows the name, because every folio reader keys
 * plain objects by name and refuses it too; a part this accepts is one they
 * can read.
 */
export const wellFormedXmlIssue = (
  xml: string,
  undeclaredEntities: XmlUndeclaredEntities = "refuse",
): string | null => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "verdict",
    conformance: "well-formed",
    references: "raw",
    undeclaredEntities,
  });
  return parsed.status === "parsed" ? null : parsed.reason;
};

/**
 * @param inheritedNamespaceScope the bindings in scope around `xml`, for a
 *   fragment captured out of a part whose root declared them.
 */
export const parseStreamingXml = (
  xml: string,
  inheritedNamespaceScope: XmlNamespaceScope = EMPTY_NAMESPACE_SCOPE,
): ParseXmlResult => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "tree",
    inheritedNamespaceScope,
  });
  return parsed.status === "parsed"
    ? { status: "parsed", value: parsed.value }
    : { status: "unsupported" };
};

/** Retain source spans for every element without changing the parsed namespace tree. */
export const parseStreamingXmlWithSourceRanges = (
  xml: string,
  inheritedNamespaceScope: XmlNamespaceScope = EMPTY_NAMESPACE_SCOPE,
): ParseXmlResult => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "tree",
    inheritedNamespaceScope,
    sourceRanges: "tracked",
  });
  return parsed.status === "parsed"
    ? { status: "parsed", value: parsed.value }
    : { status: "unsupported" };
};

/** Retain the body tree while collecting identities for import repair. */
export const parseStreamingXmlWithIdentityVisitor = (
  xml: string,
  visitOpenTag: OpenTagScanVisitor,
): ParseXmlResult => {
  const parsed = parseStreamingXmlInternal({ xml, mode: "tree", visitOpenTag });
  return parsed.status === "parsed"
    ? { status: "parsed", value: parsed.value }
    : { status: "unsupported" };
};

/** Rewrite selected decimal attribute values while preserving every other source byte. */
export const rewriteStreamingXmlDecimalAttributes = (
  xml: string,
  visitOpenTag: OpenTagVisitor,
): { status: "rewritten"; value: string } | { status: "unsupported" } => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "attributes",
    visitOpenTag,
  });
  if (parsed.status === "unsupported") {
    return { status: "unsupported" };
  }
  const chunks: string[] = [];
  let cursor = 0;
  for (const replacement of parsed.replacements.toSorted(
    (left, right) => left.start - right.start,
  )) {
    if (!/^\d+$/u.test(replacement.value)) {
      return { status: "unsupported" };
    }
    chunks.push(xml.slice(cursor, replacement.start), replacement.value);
    cursor = replacement.end;
  }
  chunks.push(xml.slice(cursor));
  return { status: "rewritten", value: chunks.join("") };
};

/** Visit identity attributes and namespace bindings without building an XML tree. */
export const scanStreamingXmlNumericIdAttributes = (
  xml: string,
  visitOpenTag: OpenTagScanVisitor,
):
  | {
      status: "unsupported";
    }
  | {
      status: "scanned";
    } => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "attributes",
    visitOpenTag,
    retainAttribute: ({ attributeName, elementName }) => {
      if (attributeName === "xmlns" || attributeName.startsWith("xmlns:"))
        return true;
      return isOoxmlNumericIdAttributeName({ elementName, attributeName });
    },
  });
  return parsed.status === "parsed"
    ? { status: "scanned" as const }
    : { status: "unsupported" };
};

/** Visit namespace-resolved start tags and source spans without building a tree. */
export const scanStreamingXmlElements = (
  xml: string,
  visitElement: ElementScanVisitor,
):
  | {
      status: "unsupported";
    }
  | {
      status: "scanned";
    } => {
  const parsed = parseStreamingXmlInternal({
    xml,
    mode: "attributes",
    visitOpenTag: () => null,
    visitElement,
  });
  return parsed.status === "parsed"
    ? { status: "scanned" as const }
    : { status: "unsupported" };
};

type ParsedOpenTag =
  | {
      status: "parsed";
      element: XmlElement;
      name: string;
      nameEnd: number;
      selfClosing: boolean;
      attributeValueSpans: ReadonlyMap<string, AttributeValueSpan> | undefined;
    }
  | Rejected;

type ParseOpenTagOptions = {
  xml: string;
  start: number;
  close: number;
  spanMode: "none" | "all" | "identities";
  rules: ReadRules;
  retainAttribute?: ((options: AttributeNameOptions) => boolean) | undefined;
};

const parseOpenTag = ({
  xml,
  start,
  close,
  spanMode,
  rules,
  retainAttribute,
}: ParseOpenTagOptions): ParsedOpenTag => {
  const { validNames } = rules;
  if (validNames !== undefined && isXmlWhitespace(xml.charCodeAt(start))) {
    return reject("whitespace before a start-tag name");
  }
  let cursor = skipWhitespace(xml, start, close);
  const nameStart = cursor;
  cursor = scanName(xml, cursor, close);
  if (cursor === nameStart) {
    return reject("a start tag without a name");
  }

  const name = xml.slice(nameStart, cursor);
  const nameEnd = cursor;
  if (isUnsafePropertyName(name)) {
    return reject(`the reserved name ${name}`);
  }
  if (validNames !== undefined && !isValidName(name, validNames)) {
    return reject(`the invalid element name ${name}`);
  }
  let attributes: Record<string, string> | undefined;
  let attributeValueSpans: Map<string, AttributeValueSpan> | undefined;
  let selfClosing = false;

  while (cursor < close) {
    const separated = cursor;
    cursor = skipWhitespace(xml, cursor, close);
    if (cursor >= close) {
      break;
    }
    if (xml.charCodeAt(cursor) === 47) {
      selfClosing = true;
      if (validNames !== undefined && cursor + 1 !== close) {
        return reject("characters between / and > in an empty-element tag");
      }
      cursor = skipWhitespace(xml, cursor + 1, close);
      if (cursor !== close) {
        return reject("characters after / in a start tag");
      }
      break;
    }
    if (validNames !== undefined && cursor === separated) {
      return reject("attributes without whitespace between them");
    }

    const attributeNameStart = cursor;
    cursor = scanName(xml, cursor, close);
    if (cursor === attributeNameStart) {
      return reject("an attribute without a name");
    }
    const attributeName = xml.slice(attributeNameStart, cursor);
    if (isUnsafePropertyName(attributeName)) {
      return reject(`the reserved name ${attributeName}`);
    }
    if (validNames !== undefined && !isValidName(attributeName, validNames)) {
      return reject(`the invalid attribute name ${attributeName}`);
    }
    cursor = skipWhitespace(xml, cursor, close);
    if (xml.charCodeAt(cursor) !== 61) {
      return reject(`the attribute ${attributeName} without a value`);
    }
    cursor = skipWhitespace(xml, cursor + 1, close);

    const quote = xml.charCodeAt(cursor);
    if (quote !== 34 && quote !== 39) {
      return reject(`the unquoted attribute ${attributeName}`);
    }
    const valueStart = cursor + 1;
    cursor = xml.indexOf(quote === 34 ? '"' : "'", valueStart);
    if (cursor === -1 || cursor > close) {
      return reject(`the unterminated attribute ${attributeName}`);
    }
    const raw = xml.slice(valueStart, cursor);
    if (validNames !== undefined && raw.includes("<")) {
      return reject(`a < in the attribute ${attributeName}`);
    }
    const decoded = rules.readValue(raw);
    if (decoded === null) {
      return reject(
        `a reference the reader does not accept in ${attributeName}`,
      );
    }
    if (
      retainAttribute === undefined ||
      retainAttribute({ attributeName, elementName: name })
    ) {
      attributes ??= {};
      if (
        validNames !== undefined &&
        Object.hasOwn(attributes, attributeName)
      ) {
        return reject(`the repeated attribute ${attributeName}`);
      }
      attributes[attributeName] = decoded;
      if (
        spanMode === "all" ||
        (spanMode === "identities" &&
          isOoxmlNumericIdAttributeName({ elementName: name, attributeName }))
      ) {
        attributeValueSpans ??= new Map();
        attributeValueSpans.set(attributeName, {
          start: valueStart,
          end: cursor,
        });
      }
    }
    cursor += 1;
  }

  const element: XmlElement = { type: "element", name };
  if (attributes) {
    element.attributes = attributes;
  }
  return {
    status: "parsed",
    element,
    name,
    nameEnd,
    selfClosing,
    attributeValueSpans,
  };
};

const appendElement = (parent: XmlElement, child: XmlElement): void => {
  parent.elements ??= [];
  parent.elements.push(child);
};

const appendText = (
  text: string,
  stack: ElementFrame[],
  mergeAdjacentText: boolean,
  rules: ReadRules,
): boolean => {
  if (text.length === 0) {
    return true;
  }
  const decoded = rules.readValue(text);
  if (decoded === null) {
    return false;
  }
  const parent = stack.at(-1)?.element;
  const previous = mergeAdjacentText ? parent?.elements?.at(-1) : undefined;
  if (previous?.type === "text" && typeof previous.text === "string") {
    previous.text += decoded;
    return true;
  }
  return appendRawText(decoded, stack, rules);
};

const appendRawText = (
  text: string,
  stack: ElementFrame[],
  rules: ReadRules,
): boolean => {
  const parent = stack.at(-1)?.element;
  if (!parent) {
    return isXmlWhitespaceText(text) || rules.conformance === "tolerant";
  }
  appendElement(parent, { type: "text", text });
  return true;
};

type ProcessingInstruction = { target: string; body: string };

const readInstruction = (
  xml: string,
  open: number,
  close: number,
): ProcessingInstruction => {
  let targetEnd = open + 2;
  while (targetEnd < close && !isXmlWhitespace(xml.charCodeAt(targetEnd))) {
    targetEnd += 1;
  }
  return {
    target: xml.slice(open + 2, targetEnd),
    body: xml.slice(targetEnd, close).trim(),
  };
};

type InstructionIssueOptions = {
  instruction: ProcessingInstruction;
  open: number;
  documentStart: number;
  rules: ReadRules;
};

/** XML 1.0 §2.8: a declaration opens with `version="1.<digits>"`. */
const XML_DECLARATION_VERSION = /^version\s*=\s*(["'])1\.\d+\1(?:\s|$)/u;

const instructionIssue = ({
  instruction: { target, body },
  open,
  documentStart,
  rules,
}: InstructionIssueOptions): string | null => {
  if (!isValidName(target, rules.validNames ?? new Set())) {
    return "a processing instruction without a valid target";
  }
  if (target.toLowerCase() !== "xml") {
    return null;
  }
  if (target !== "xml" || open !== documentStart) {
    return "an XML declaration after the start of the document";
  }
  return XML_DECLARATION_VERSION.test(body)
    ? null
    : "an XML declaration without a version";
};

const findTagClose = (xml: string, start: number): number => {
  let quote = 0;
  for (let cursor = start; cursor < xml.length; cursor += 1) {
    const code = xml.charCodeAt(cursor);
    if (quote !== 0) {
      if (code === quote) {
        quote = 0;
      }
      continue;
    }
    if (code === 34 || code === 39) {
      quote = code;
      continue;
    }
    if (code === 62) {
      return cursor;
    }
  }
  return -1;
};

const isXmlWhitespace = (code: number): boolean =>
  code === 9 || code === 10 || code === 13 || code === 32;

/**
 * Whether `text` is only the whitespace XML 1.0 admits outside the root:
 * space, tab, CR and LF. `String.prototype.trim` also strips U+00A0 and
 * other Unicode spaces, which XML does not.
 */
const isXmlWhitespaceText = (text: string): boolean => {
  for (let index = 0; index < text.length; index += 1) {
    if (!isXmlWhitespace(text.charCodeAt(index))) return false;
  }
  return true;
};

const skipWhitespace = (xml: string, start: number, limit: number): number => {
  let cursor = start;
  while (cursor < limit && isXmlWhitespace(xml.charCodeAt(cursor))) {
    cursor += 1;
  }
  return cursor;
};

const scanName = (xml: string, start: number, limit: number): number => {
  let cursor = start;
  while (cursor < limit) {
    const code = xml.charCodeAt(cursor);
    if (isXmlWhitespace(code) || code === 47 || code === 61 || code === 62) {
      break;
    }
    cursor += 1;
  }
  return cursor;
};

/** XML 1.0 (Fifth Edition) §2.3 `Name`. */
const NAME_START_CHARACTER = String.raw`:A-Z_a-z\u{C0}-\u{D6}\u{D8}-\u{F6}\u{F8}-\u{2FF}\u{370}-\u{37D}\u{37F}-\u{1FFF}\u{200C}-\u{200D}\u{2070}-\u{218F}\u{2C00}-\u{2FEF}\u{3001}-\u{D7FF}\u{F900}-\u{FDCF}\u{FDF0}-\u{FFFD}\u{10000}-\u{EFFFF}`;
const NAME_CHARACTER = String.raw`${NAME_START_CHARACTER}\-.0-9\u{B7}\u{300}-\u{36F}\u{203F}-\u{2040}`;
const XML_NAME = new RegExp(
  `^[${NAME_START_CHARACTER}][${NAME_CHARACTER}]*$`,
  "u",
);

const isValidName = (name: string, validNames: Set<string>): boolean => {
  if (validNames.has(name)) return true;
  if (!XML_NAME.test(name)) return false;
  validNames.add(name);
  return true;
};

/** XML 1.0 §2.2 `Char`. */
const isXmlCharacter = (codePoint: number): boolean =>
  codePoint === 0x09 ||
  codePoint === 0x0a ||
  codePoint === 0x0d ||
  (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
  (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
  (codePoint >= 0x10000 && codePoint <= 0x10ffff);

const containsOnlyXmlCharacters = (xml: string): boolean => {
  for (let index = 0; index < xml.length; index += 1) {
    const code = xml.charCodeAt(index);
    if (code >= 0x20 && code < 0xd800) continue;
    if (code === 0x09 || code === 0x0a || code === 0x0d) continue;
    if (code >= 0xe000 && code <= 0xfffd) continue;
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = xml.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        index += 1;
        continue;
      }
    }
    return false;
  }
  return true;
};

type ReferenceDecoder = (entity: string) => string | null;

/**
 * Decode every reference in `value`; `null` when one is refused.
 *
 * The decoder answers for the text between `&` and `;`, and a `null` there
 * refuses the whole value.
 */
const decodeReferences = (
  value: string,
  decodeReference: ReferenceDecoder,
): string | null => {
  const firstEntity = value.indexOf("&");
  if (firstEntity === -1) {
    return value;
  }

  const parts: string[] = [];
  let cursor = 0;
  let entityStart = firstEntity;
  while (entityStart !== -1) {
    parts.push(value.slice(cursor, entityStart));
    const entityEnd = value.indexOf(";", entityStart + 1);
    if (entityEnd === -1) {
      return null;
    }
    const decoded = decodeReference(value.slice(entityStart + 1, entityEnd));
    if (decoded === null) {
      return null;
    }
    parts.push(decoded);
    cursor = entityEnd + 1;
    entityStart = value.indexOf("&", cursor);
  }
  parts.push(value.slice(cursor));
  return parts.join("");
};

const normalizeLineEndings = (value: string): string => {
  if (!value.includes("\r")) {
    return value;
  }
  return value.replaceAll("\r\n", "\n").replaceAll("\r", "\n");
};

const builtInEntity = (entity: string): string | null => {
  switch (entity) {
    case "amp":
      return BUILT_IN_ENTITIES.amp;
    case "apos":
      return BUILT_IN_ENTITIES.apos;
    case "gt":
      return BUILT_IN_ENTITIES.gt;
    case "lt":
      return BUILT_IN_ENTITIES.lt;
    case "quot":
      return BUILT_IN_ENTITIES.quot;
    default:
      return null;
  }
};

type CharacterReference = { codePoint: number; hexMarker: "x" | "X" | null };

const parseCharacterReference = (entity: string): CharacterReference | null => {
  if (entity.charCodeAt(0) !== 35) {
    return null;
  }
  const marker = entity.charAt(1);
  const hexMarker = marker === "x" || marker === "X" ? marker : null;
  const digits = entity.slice(hexMarker === null ? 1 : 2);
  if (digits.length === 0 || !hasOnlyDigits(digits, hexMarker !== null)) {
    return null;
  }
  return {
    codePoint: Number.parseInt(digits, hexMarker === null ? 10 : 16),
    hexMarker,
  };
};

/** The `ooxml` reader's references: the built-ins and any scalar value. */
const decodeOoxmlReference: ReferenceDecoder = (entity) => {
  const builtIn = builtInEntity(entity);
  if (builtIn !== null) return builtIn;
  const reference = parseCharacterReference(entity);
  if (reference === null) return null;
  const { codePoint } = reference;
  if (
    !Number.isInteger(codePoint) ||
    codePoint <= 0 ||
    codePoint > 0x10ffff ||
    (codePoint >= 0xd800 && codePoint <= 0xdfff)
  ) {
    return null;
  }
  return String.fromCodePoint(codePoint);
};

/** XML 1.0 §4.1: `&#x` in lower case, and only characters §2.2 admits. */
const decodeWellFormedReference: ReferenceDecoder = (entity) => {
  const builtIn = builtInEntity(entity);
  if (builtIn !== null) return builtIn;
  const reference = parseCharacterReference(entity);
  if (
    reference === null ||
    reference.hexMarker === "X" ||
    !Number.isInteger(reference.codePoint) ||
    !isXmlCharacter(reference.codePoint)
  ) {
    return null;
  }
  return String.fromCodePoint(reference.codePoint);
};

/**
 * A reference that is syntactically one: decoded when XML defines it, kept as
 * written when it names an entity no declaration defines.
 */
const decodeReferenceKeepingUndeclared: ReferenceDecoder = (entity) => {
  const decoded = decodeWellFormedReference(entity);
  if (decoded !== null) return decoded;
  if (entity.charCodeAt(0) === 35) {
    const reference = parseCharacterReference(entity);
    return reference === null || reference.hexMarker === "X"
      ? null
      : `&${entity};`;
  }
  return XML_NAME.test(entity) ? `&${entity};` : null;
};

/**
 * Names earlier folio releases decoded in DOCX text, beyond XML's five. A part
 * that used one opened with the character it names, and still does.
 */
const LEGACY_NAMED_REFERENCES: ReadonlyMap<string, string> = new Map([
  ["nbsp", String.fromCodePoint(0xa0)],
  ["copy", String.fromCodePoint(0xa9)],
  ["reg", String.fromCodePoint(0xae)],
  ["trade", String.fromCodePoint(0x2122)],
  ["mdash", String.fromCodePoint(0x2014)],
  ["ndash", String.fromCodePoint(0x2013)],
  ["hellip", String.fromCodePoint(0x2026)],
  ["laquo", String.fromCodePoint(0xab)],
  ["raquo", String.fromCodePoint(0xbb)],
  ["lsquo", String.fromCodePoint(0x2018)],
  ["rsquo", String.fromCodePoint(0x2019)],
  ["ldquo", String.fromCodePoint(0x201c)],
  ["rdquo", String.fromCodePoint(0x201d)],
  ["bull", String.fromCodePoint(0x2022)],
  ["para", String.fromCodePoint(0xb6)],
  ["sect", String.fromCodePoint(0xa7)],
  ["deg", String.fromCodePoint(0xb0)],
  ["frac12", String.fromCodePoint(0xbd)],
  ["frac14", String.fromCodePoint(0xbc)],
  ["frac34", String.fromCodePoint(0xbe)],
  ["cent", String.fromCodePoint(0xa2)],
  ["pound", String.fromCodePoint(0xa3)],
  ["curren", String.fromCodePoint(0xa4)],
  ["yen", String.fromCodePoint(0xa5)],
  ["euro", String.fromCodePoint(0x20ac)],
  ["dollar", String.fromCodePoint(0x24)],
  ["fnof", String.fromCodePoint(0x192)],
  ["inr", String.fromCodePoint(0x20b9)],
  ["af", String.fromCodePoint(0x60b)],
  ["birr", String.fromCodePoint(0x1265, 0x122d)],
  ["peso", String.fromCodePoint(0x20b1)],
  ["rub", String.fromCodePoint(0x20bd)],
  ["won", String.fromCodePoint(0x20a9)],
  ["yuan", String.fromCodePoint(0xa5)],
  ["cedil", String.fromCodePoint(0xb8)],
]);

/** The longest reference body earlier releases looked for a `;` within. */
const LEGACY_REFERENCE_WINDOW = 32;

/**
 * A reference the `ooxml` reader refuses, read the way earlier releases read
 * it: a character reference to a control character, a surrogate or U+0000
 * reads as nothing, one beyond U+10FFFF and an unknown name stay as written,
 * and the digits of a character reference are read up to the first that is
 * not one.
 */
const legacyReference = (entity: string): string | null => {
  const named = LEGACY_NAMED_REFERENCES.get(entity);
  if (named !== undefined) return named;
  if (entity.charCodeAt(0) !== 35) return null;
  const marker = entity.charAt(1);
  const hexadecimal = marker === "x" || marker === "X";
  const codePoint = Number.parseInt(
    entity.slice(hexadecimal ? 2 : 1),
    hexadecimal ? 16 : 10,
  );
  if (Number.isNaN(codePoint) || codePoint < 0 || codePoint > 0x10ffff) {
    return null;
  }
  if (
    codePoint === 0 ||
    (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
    (codePoint < 0x20 &&
      codePoint !== 0x09 &&
      codePoint !== 0x0a &&
      codePoint !== 0x0d)
  ) {
    return "";
  }
  return String.fromCodePoint(codePoint);
};

const decodeTolerantReferences = (value: string): string => {
  const parts: string[] = [];
  let literalStart = 0;
  let cursor = value.indexOf("&");
  while (cursor !== -1) {
    let end = cursor + 1;
    while (
      end < value.length &&
      value.charCodeAt(end) !== 59 &&
      end - cursor <= LEGACY_REFERENCE_WINDOW
    ) {
      end += 1;
    }
    if (end >= value.length || value.charCodeAt(end) !== 59) {
      cursor = value.indexOf("&", cursor + 1);
      continue;
    }
    const entity = value.slice(cursor + 1, end);
    const decoded =
      entity.length === 0
        ? null
        : (decodeOoxmlReference(entity) ?? legacyReference(entity));
    if (decoded === null) {
      cursor = value.indexOf("&", cursor + 1);
      continue;
    }
    parts.push(value.slice(literalStart, cursor), decoded);
    literalStart = end + 1;
    cursor = value.indexOf("&", literalStart);
  }
  if (parts.length === 0) {
    return value;
  }
  parts.push(value.slice(literalStart));
  return parts.join("");
};

const hasOnlyDigits = (value: string, hexadecimal: boolean): boolean => {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    const decimal = code >= 48 && code <= 57;
    const lowerHex = hexadecimal && code >= 97 && code <= 102;
    const upperHex = hexadecimal && code >= 65 && code <= 70;
    if (!decimal && !lowerHex && !upperHex) {
      return false;
    }
  }
  return true;
};

const isUnsafePropertyName = (name: string): boolean => {
  switch (name) {
    case "__defineGetter__":
    case "__defineSetter__":
    case "__lookupGetter__":
    case "__lookupSetter__":
    case "__proto__":
    case "constructor":
    case "hasOwnProperty":
    case "prototype":
    case "toString":
    case "valueOf":
      return true;
    default:
      return false;
  }
};
