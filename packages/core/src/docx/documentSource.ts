import type { BlockContent, DocumentBody } from "../types/document";
import { requiresXmlSpacePreserve } from "@stll/docx-core";
import { canonicalJson } from "../utils/canonicalJson";
import { getXmlSourceRange } from "./streamingXmlParser";
import { getParagraphPropertySourceCandidate } from "./paragraphPropertySource";
import {
  getAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  getTextContent,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";
import { XML_NAMESPACE_URI } from "./xmlNamespaceContext";
import { isSafeCapturedXmlDocument, isSafeParsedXmlDocument } from "./verbatimCapture";
import { TaggedError } from "better-result";

class DocumentSourceReplayError extends TaggedError("DocumentSourceReplayError")<{
  message: string;
}> {}

const validatedSources = new WeakMap<NonNullable<DocumentBody["source"]>, string>();

const safeSourceXml = (xml: string): string => {
  if (!isSafeCapturedXmlDocument(xml)) {
    throw new DocumentSourceReplayError({ message: "Document source XML is not safe to replay." });
  }
  return xml;
};

const safeCapturedSource = (source: NonNullable<DocumentBody["source"]>): string => {
  if (validatedSources.get(source) === source.xml) return source.xml;
  const xml = safeSourceXml(source.xml);
  validatedSources.set(source, xml);
  return xml;
};

const shellFingerprint = (body: DocumentBody): string =>
  canonicalJson({
    background: body.background,
    finalSectionProperties: body.finalSectionProperties,
  });

// The compact value snapshot is exact, not a hash. Keeping strings rather than
// cloned model subgraphs lets graph-preserving document clones share baselines.
const groupFingerprint = (content: BlockContent[]): string => JSON.stringify(content);
const XML_NAMESPACE_URIS = new Set([XML_NAMESPACE_URI]);
const RELATIONSHIP_REFERENCE_ATTRIBUTES = new Set(["id", "embed", "link"]);

/** Model-neutral parser repairs still require the writer to replace invalid source XML. */
const requiresSourceRepair = (element: XmlElement, excluded?: ReadonlySet<XmlElement>): boolean => {
  if (excluded?.has(element)) return false;
  for (const attribute in element.attributes) {
    if (
      element.attributes?.[attribute] === "" &&
      RELATIONSHIP_REFERENCE_ATTRIBUTES.has(getLocalName(attribute)) &&
      OFFICE_RELATIONSHIP_NAMESPACE_URIS.has(resolveAttributeNamespaceUri(element, attribute) ?? "")
    ) {
      return true;
    }
  }
  const localName = getLocalName(element.name ?? "");
  if (
    WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "") &&
    (localName === "t" || localName === "delText") &&
    requiresXmlSpacePreserve(getTextContent(element)) &&
    getAttributeByNamespaceUri(element, XML_NAMESPACE_URIS, "space") !== "preserve"
  ) {
    return true;
  }
  return element.elements?.some((child) => requiresSourceRepair(child, excluded)) ?? false;
};

/** Capture before package normalizers; a changed model must invalidate replay. */
export const captureDocumentSource = (
  body: DocumentBody,
  source: { xml: string; root: XmlElement; groups: ReadonlyMap<XmlElement, BlockContent[]> },
): void => {
  // Malformed packages keep the existing rebuilding path; they cannot grant
  // raw replay authority merely because the tolerant parser accepted them.
  if (!isSafeParsedXmlDocument(source.xml, source.root)) return;
  const blocks: NonNullable<DocumentBody["source"]>["blocks"] = new Map();
  const capturedGroups = new Set<XmlElement>();
  for (const [element, content] of source.groups) {
    const range = getXmlSourceRange(element);
    if (!range || content.length === 0) continue;
    capturedGroups.add(element);
    // Mutable model records cannot serve as their own baseline: callers may
    // edit before the first save. Capture compact values, never a deep clone.
    const record = {
      ...range,
      content,
      fingerprint: requiresSourceRepair(element) ? null : groupFingerprint(content),
    };
    for (const block of content) blocks.set(block, record);
  }
  body.source = {
    xml: source.xml,
    shellFingerprint: requiresSourceRepair(source.root, capturedGroups)
      ? null
      : shellFingerprint(body),
    blocks,
  };
  validatedSources.set(body.source, source.xml);
};

type ReplayDocumentSourceOptions = {
  body: DocumentBody;
  serialize: (blocks: BlockContent[]) => string;
};

/**
 * Source ownership is by model identity, including through structuredClone.
 * A model edit invalidates only its source group. Structural edits and shell
 * edits take the rebuilding path until their source placement is represented.
 */
export const replayDocumentSource = ({
  body,
  serialize,
}: ReplayDocumentSourceOptions): string | null => {
  const source = body.source;
  if (!source) return null;
  if (source.shellFingerprint !== shellFingerprint(body)) return null;

  const replacements: { start: number; end: number; xml: string }[] = [];
  let index = 0;
  let previousEnd = 0;
  while (index < body.content.length) {
    const block = body.content.at(index);
    const owner =
      block?.type === "paragraph" ? (getParagraphPropertySourceCandidate(block) ?? block) : block;
    const record = owner && source.blocks.get(owner);
    if (!record || record.start < previousEnd) return null;
    const content = body.content.slice(index, index + record.content.length);
    if (
      content.length !== record.content.length ||
      content.some((item, offset) => {
        const candidate =
          item.type === "paragraph" ? (getParagraphPropertySourceCandidate(item) ?? item) : item;
        return candidate !== record.content.at(offset);
      })
    )
      return null;
    if (groupFingerprint(content) !== record.fingerprint) {
      replacements.push({ start: record.start, end: record.end, xml: serialize(content) });
    }
    previousEnd = record.end;
    index += content.length;
  }
  // Deleting a tail must not replay the source's deleted blocks.
  if (source.blocks.size !== body.content.length) return null;
  if (replacements.length === 0) return safeCapturedSource(source);
  const chunks: string[] = [];
  let cursor = 0;
  for (const replacement of replacements) {
    chunks.push(source.xml.slice(cursor, replacement.start), replacement.xml);
    cursor = replacement.end;
  }
  chunks.push(source.xml.slice(cursor));
  return safeSourceXml(chunks.join(""));
};
