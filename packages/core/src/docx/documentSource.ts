import type { BlockContent, Document, DocumentBody } from "../types/document";
import { cloneDocumentWithParagraphPropertySources } from "./paragraphPropertySource";
import { hasUnsynthesizedReplyRanges } from "./commentReplyMarkers";
import { requiresXmlSpacePreserve } from "@stll/docx-core";
import {
  getSourceReplayToken,
  registerSourceReplayDocument,
  type SourceReplayToken,
} from "@stll/docx-core/ops";
import { getXmlSourceRange } from "./streamingXmlParser";
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
import { isSafeCapturedXmlDocument, getSingleParsedXmlDocumentElement } from "./verbatimCapture";
import { TaggedError } from "better-result";

class DocumentSourceReplayError extends TaggedError("DocumentSourceReplayError")<{
  message: string;
}> {}

type SourceGroup = {
  source: SourcePart;
  start: number;
  end: number;
  content: readonly BlockContent[];
  repair: "required" | "none";
};

type SourcePart = {
  xml: string;
  content: readonly BlockContent[];
  background: DocumentBody["background"];
  finalSectionProperties: DocumentBody["finalSectionProperties"];
  repair: "required" | "none";
};

const blockSources = new WeakMap<BlockContent, SourceGroup>();
const bodySources = new WeakMap<DocumentBody, SourcePart>();
const tokenSources = new WeakMap<SourceReplayToken, SourcePart>();
const validatedSources = new WeakSet<SourcePart>();

const safeSourceXml = (xml: string): string => {
  if (!isSafeCapturedXmlDocument(xml)) {
    throw new DocumentSourceReplayError({ message: "Document source XML is not safe to replay." });
  }
  return xml;
};

/** Bind replay only after import repairs, before handing an immutable model to an editor. */
export const trackDocumentSource = (document: Document): void => {
  const source = bodySources.get(document.package.document);
  const token = registerSourceReplayDocument(document);
  if (source) tokenSources.set(token, source);
};

/** Import repairs cannot retain authority for the model they changed in place. */
export const discardDocumentSource = (body: DocumentBody): void => {
  bodySources.delete(body);
  delete body.source;
};

export const documentSourceXml = (
  document: Document,
  token: SourceReplayToken,
): string | undefined =>
  getSourceReplayToken(document) === token ? tokenSources.get(token)?.xml : undefined;

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

/** Retain structural references only; replay is explicitly enabled by a tracked save. */
export const captureDocumentSource = (
  body: DocumentBody,
  source: { xml: string; root: XmlElement; groups: ReadonlyMap<XmlElement, BlockContent[]> },
): void => {
  const root = getSingleParsedXmlDocumentElement(source.root);
  if (!root || !getXmlSourceRange(root)) return;
  const capturedGroups = new Set<XmlElement>();
  for (const [element, content] of source.groups) {
    if (getXmlSourceRange(element) && content.length > 0) capturedGroups.add(element);
  }
  const part = {
    xml: source.xml,
    content: body.content,
    background: body.background,
    finalSectionProperties: body.finalSectionProperties,
    repair: requiresSourceRepair(source.root, capturedGroups) ? "required" : "none",
  } satisfies SourcePart;
  for (const [element, content] of source.groups) {
    const range = getXmlSourceRange(element);
    if (!range || content.length === 0) continue;
    const record = {
      source: part,
      ...range,
      content,
      repair: requiresSourceRepair(element) ? "required" : "none",
    } as const satisfies SourceGroup;
    for (const block of content) blockSources.set(block, record);
  }
  body.source = { xml: source.xml };
  bodySources.set(body, part);
};

type ReplayDocumentSourceOptions = {
  document: Document;
  token: SourceReplayToken | undefined;
  serialize: (blocks: BlockContent[]) => string;
};

/** New blocks serialize from the model; unchanged immutable identities replay their source slice. */
export const replayDocumentSource = ({
  document,
  token,
  serialize,
}: ReplayDocumentSourceOptions): string | null => {
  if (token === undefined || getSourceReplayToken(document) !== token) return null;
  const source = tokenSources.get(token);
  const body = document.package.document;
  if (
    !source ||
    source.repair === "required" ||
    source.background !== body.background ||
    source.finalSectionProperties !== body.finalSectionProperties ||
    source.content.length !== body.content.length
  )
    return null;
  if (!validatedSources.has(source)) {
    if (!isSafeCapturedXmlDocument(source.xml)) return null;
    validatedSources.add(source);
  }
  const replacements: { start: number; end: number; xml: string }[] = [];
  let index = 0;
  let previousEnd = 0;
  while (index < source.content.length) {
    const original = source.content.at(index);
    const record = original && blockSources.get(original);
    if (!record || record.start < previousEnd) return null;
    let changed = record.repair === "required";
    for (let offset = 0; offset < record.content.length; offset += 1) {
      const block = body.content.at(index + offset);
      if (!block) return null;
      const retained = blockSources.get(block);
      // Moving an authored identity changes source placement; rebuild rather
      // than replaying XML from its previous position.
      if (
        retained &&
        (retained.source !== source || retained !== record || block !== record.content.at(offset))
      )
        return null;
      changed ||= retained === undefined;
    }
    if (changed) {
      replacements.push({
        start: record.start,
        end: record.end,
        xml: serialize(body.content.slice(index, index + record.content.length)),
      });
    }
    previousEnd = record.end;
    index += record.content.length;
  }
  if (replacements.length === 0) return source.xml;
  const chunks: string[] = [];
  let cursor = 0;
  for (const replacement of replacements) {
    chunks.push(source.xml.slice(cursor, replacement.start), replacement.xml);
    cursor = replacement.end;
  }
  chunks.push(source.xml.slice(cursor));
  return safeSourceXml(chunks.join(""));
};

/** Export normalizers may write only into model-derived blocks, never retained source identities. */
export const prepareSourceReplayExport = (document: Document): Document => {
  const token = getSourceReplayToken(document);
  if (token === undefined) return document;
  const source = tokenSources.get(token);
  const repliesNeedMarkers = hasUnsynthesizedReplyRanges(document);
  let changed = false;
  const content = document.package.document.content.map((block) => {
    const retained = source !== undefined && blockSources.get(block)?.source === source;
    if (retained && !repliesNeedMarkers) return block;
    changed = true;
    const cloned = cloneDocumentWithParagraphPropertySources({
      package: { document: { content: [block] } },
    }).package.document.content.at(0);
    if (!cloned) throw new DocumentSourceReplayError({ message: "Export clone lost its block." });
    return cloned;
  });
  if (!changed) return document;
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content },
    },
  };
};
