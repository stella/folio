/** Conservative body-paragraph splices; source offsets never come from reserialization. */
import { canonicalJson } from "../utils/canonicalJson";
import { parseDocumentBody } from "./documentParser";
import { paraIdAttribute } from "./paraIdAttribute";
import { spliceXml, type XmlSplice } from "./selectiveXmlPatch";
import { readRootNamespaceBindings, serializePartElement } from "./serializer/partNamespaces";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  NAMESPACES,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const isWordElement = (element: XmlElement, name: string): boolean =>
  getLocalName(element.name) === name &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "");

type BodyBlock = { element: XmlElement; start: number; end: number };

/** Match XML tokens, skipping quoted delimiters, comments, CDATA and processing instructions. */
const XML_TOKEN = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;

const readBody = (xml: string) => {
  const root = parseXmlDocument(xml);
  if (!root || !isWordElement(root, "document")) return null;
  // Strict conversion belongs to full repack; a Transitional fragment cannot
  // inherit a Strict namespace, even when the producer spells its prefix `w`.
  if (getNamespaceUri(root) !== NAMESPACES.w) return null;
  const body = getChildElements(root).find((child) => isWordElement(child, "body"));
  if (!body) return null;
  const children = getChildElements(body);
  const blocks: BodyBlock[] = [];
  const stack: string[] = [];
  let bodyDepth = -1;
  let bodyEnd = -1;
  let start = -1;
  for (const token of xml.matchAll(XML_TOKEN)) {
    const tag = token[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    const closing = tag.startsWith("</");
    const name = tag.slice(closing ? 2 : 1).split(/[\s/>]/u).at(0);
    if (!name) return null;
    if (closing) {
      if (stack.pop() !== name) return null;
      if (stack.length === bodyDepth) {
        const element = children[blocks.length];
        if (!element || start < 0) return null;
        blocks.push({ element, start, end: token.index + tag.length });
        start = -1;
      }
      if (name === body.name && stack.length === bodyDepth - 1) bodyEnd = token.index;
      continue;
    }
    if (name === body.name && stack.length === 1) bodyDepth = 2;
    if (stack.length === bodyDepth) {
      if (children[blocks.length]?.name !== name) return null;
      start = token.index;
      if (tag.endsWith("/>")) {
        const element = children[blocks.length];
        if (!element) return null;
        blocks.push({ element, start, end: token.index + tag.length });
        start = -1;
      }
    }
    if (!tag.endsWith("/>")) stack.push(name);
  }
  if (stack.length > 0 || blocks.length !== children.length || bodyEnd < 0) return null;
  return { root, blocks, bodyEnd };
};

function* descendants(element: XmlElement): Generator<XmlElement> {
  yield element;
  for (const child of getChildElements(element)) yield* descendants(child);
}

/** Cross-paragraph ranges require a wider edit contract than paragraph identity. */
const RANGE_ELEMENTS = new Set([
  "commentRangeStart", "commentRangeEnd", "commentReference",
  "bookmarkStart", "bookmarkEnd", "permStart", "permEnd", "fldChar",
  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
]);
const DEPENDENT_ELEMENTS = new Set([
  "sectPr", "footnoteReference", "endnoteReference", "drawing", "pict", "object",
]);

const paragraphIsSafe = (element: XmlElement, touched: boolean): boolean => {
  for (const child of descendants(element)) {
    if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")) continue;
    const name = getLocalName(child.name);
    if (touched && DEPENDENT_ELEMENTS.has(name)) return false;
  }
  return true;
};

const paragraphIds = (root: XmlElement): Set<string> | null => {
  const ids = new Set<string>();
  for (const element of descendants(root)) {
    if (!isWordElement(element, "p")) continue;
    const id = paraIdAttribute(element);
    if (!id || !/^[0-9a-f]{8}$/iu.test(id) || /^0+$/u.test(id) || ids.has(id.toUpperCase())) {
      return null;
    }
    ids.add(id.toUpperCase());
  }
  return ids;
};

type ParagraphFragmentOptions = {
  xml: string;
  block: BodyBlock;
  bindings: ReadonlyMap<string, string>;
};

/** Bind fragment prefixes locally: the source root may use entirely different aliases. */
const paragraphFragment = ({ xml, block, bindings }: ParagraphFragmentOptions): string => {
  const fragment = xml.slice(block.start, block.end);
  const openEnd = fragment.indexOf(">");
  const name = block.element.name ?? "w:p";
  const selfClosing = fragment[openEnd - 1] === "/";
  return serializePartElement({
    partPath: "word/document.xml",
    rootName: name,
    rootAttributes: fragment.slice(name.length + 1, selfClosing ? openEnd - 1 : openEnd).trim(),
    baselinePrefixes: [],
    sourceBindings: bindings,
    body: selfClosing ? "" : fragment.slice(openEnd + 1, fragment.lastIndexOf("</")),
  });
};

type StructuralPatchOptions = {
  originalXml: string;
  serializedXml: string;
  changedIds: ReadonlySet<string>;
};

/**
 * Insert/delete direct body paragraphs between surviving paragraph/table anchors.
 * Tables are opaque barriers: their modeled content and order must stay identical.
 * Id-less sources need an explicit ensureParaIds ingest before structural editing.
 */
export const buildStructuralDocumentPatch = ({
  originalXml, serializedXml, changedIds,
}: StructuralPatchOptions): string | null => {
  const source = readBody(originalXml);
  const current = readBody(serializedXml);
  if (!source || !current) return null;
  // A range can start in one table and end in another, enclosing a body edit
  // without placing either endpoint in an edited paragraph.
  for (const root of [source.root, current.root]) {
    for (const element of descendants(root)) {
      if (WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "") &&
        RANGE_ELEMENTS.has(getLocalName(element.name))) return null;
    }
  }
  const allSourceIds = paragraphIds(source.root);
  const allCurrentIds = paragraphIds(current.root);
  if (!allSourceIds || !allCurrentIds) return null;
  const changed = new Set([...changedIds].map((id) => id.toUpperCase()));
  for (const { element } of [...source.blocks, ...current.blocks]) {
    if (isWordElement(element, "p")) {
      const id = paraIdAttribute(element)?.toUpperCase();
      if (!id) return null;
      const touched = changed.has(id) || !allSourceIds.has(id) || !allCurrentIds.has(id);
      if (!paragraphIsSafe(element, touched)) return null;
    } else if (!isWordElement(element, "tbl") && !isWordElement(element, "sectPr")) {
      return null;
    }
  }

  // Compare both sides through the same parser, avoiding lexical differences
  // such as empty tags and attribute order. Unmodeled table bytes stay in source.
  const sourceBody = parseDocumentBody(originalXml);
  const currentBody = parseDocumentBody(serializedXml);
  const barriers = (body: typeof sourceBody) => ({
    blocks: body.content.filter((block) => block.type !== "paragraph"),
    finalSectionProperties: body.finalSectionProperties,
    background: body.background,
  });
  if (canonicalJson(barriers(sourceBody)) !== canonicalJson(barriers(currentBody))) return null;

  const indexBlocks = (blocks: readonly BodyBlock[]) => {
    let barrier = 0;
    const indexed = new Map<string, BodyBlock>();
    for (const block of blocks) {
      const key = isWordElement(block.element, "p")
        ? `p:${paraIdAttribute(block.element)?.toUpperCase()}`
        : `barrier:${barrier++}`;
      indexed.set(key, block);
    }
    return indexed;
  };
  const before = indexBlocks(source.blocks);
  const after = indexBlocks(current.blocks);
  const survivingBefore = [...before.keys()].filter((key) => after.has(key));
  const survivingAfter = [...after.keys()].filter((key) => before.has(key));
  if (canonicalJson(survivingBefore) !== canonicalJson(survivingAfter)) return null;

  const bindings = readRootNamespaceBindings(serializedXml);
  const fragmentFor = (block: BodyBlock) => paragraphFragment({ xml: serializedXml, block, bindings });
  const splices: XmlSplice[] = [];
  for (const [key, block] of before) {
    if (!after.has(key)) splices.push({ start: block.start, end: block.end, newXml: "" });
  }
  let pending: string[] = [];
  for (const [key, block] of after) {
    const original = before.get(key);
    if (!original) {
      const id = paraIdAttribute(block.element)?.toUpperCase();
      if (!id || allSourceIds.has(id)) return null;
      pending.push(fragmentFor(block));
      continue;
    }
    const id = paraIdAttribute(block.element)?.toUpperCase();
    const replacement = id !== undefined && changed.has(id)
      ? fragmentFor(block)
      : originalXml.slice(original.start, original.end);
    if (pending.length > 0 || replacement !== originalXml.slice(original.start, original.end)) {
      splices.push({ start: original.start, end: original.end, newXml: pending.join("") + replacement });
    }
    pending = [];
  }
  if (pending.length > 0) splices.push({ start: source.bodyEnd, end: source.bodyEnd, newXml: pending.join("") });
  return spliceXml(originalXml, splices);
};
