import type { DocxArchive } from "./boundedArchive";
import { loadDocxArchive } from "./boundedArchive";
import { parseRelationships, RELATIONSHIP_TYPES } from "../relsParser";
import {
  findAllDeep,
  findChild,
  findDeep,
  getAttribute,
  getAttributeAnyPrefix,
  getLocalName,
  getTextContent,
  parseXml,
  type XmlElement,
} from "../xmlParser";

const DOCUMENT_RELS_PATH = "word/_rels/document.xml.rels";

/** Document part containing an extracted paragraph. */
export type DocxParagraphSource = "header" | "body" | "footer";

/** Paragraph text and lightweight formatting metadata from a DOCX archive. */
export type ExtractedDocxParagraph = {
  index: number;
  text: string;
  source: DocxParagraphSource;
  style?: string;
  bold?: boolean;
  fontSize?: number;
  alignment?: "left" | "center" | "right" | "both";
};

/** Accepted-revision paragraph text extracted in deterministic part order. */
export type ExtractedDocxText = {
  paragraphs: ExtractedDocxParagraph[];
  charCount: number;
  view: "accepted";
};

type ParagraphProperties = Pick<ExtractedDocxParagraph, "style" | "alignment">;

type RunMetrics = {
  bold: boolean;
  fontSize?: number;
  chars: number;
};

const childElements = (element: XmlElement): XmlElement[] =>
  element.elements?.filter((child) => child.type === "element") ?? [];

const collectText = (element: XmlElement): string => {
  let text = "";

  const walk = (node: XmlElement) => {
    const localName = getLocalName(node.name ?? "");
    if (localName === "t") {
      text += getTextContent(node);
      return;
    }
    if (localName === "br") {
      text += "\n";
      return;
    }
    if (localName === "tab") {
      text += "\t";
      return;
    }
    if (localName === "del" || localName === "delText" || localName === "moveFrom") {
      return;
    }
    for (const child of childElements(node)) {
      walk(child);
    }
  };

  walk(element);
  return text;
};

const readParagraphProperties = (paragraph: XmlElement): ParagraphProperties => {
  const properties = findChild(paragraph, "w", "pPr");
  if (!properties) {
    return {};
  }

  const result: ParagraphProperties = {};
  const style = findChild(properties, "w", "pStyle");
  const styleValue = getAttributeAnyPrefix(style, "val");
  if (styleValue !== null) {
    result.style = styleValue;
  }

  const justification = findChild(properties, "w", "jc");
  const alignment = getAttributeAnyPrefix(justification, "val");
  if (
    alignment === "left" ||
    alignment === "center" ||
    alignment === "right" ||
    alignment === "both"
  ) {
    result.alignment = alignment;
  }
  return result;
};

const readRunMetrics = (paragraph: XmlElement): RunMetrics[] => {
  const metrics: RunMetrics[] = [];

  for (const run of childElements(paragraph)) {
    if (getLocalName(run.name ?? "") !== "r") {
      continue;
    }

    const properties = findChild(run, "w", "rPr");
    const boldProperty = findChild(properties, "w", "b");
    const boldValue = getAttributeAnyPrefix(boldProperty, "val");
    const bold = boldProperty !== null && boldValue !== "0" && boldValue !== "false";

    const sizeProperty = findChild(properties, "w", "sz");
    const sizeValue = getAttributeAnyPrefix(sizeProperty, "val");
    const parsedSize = sizeValue === null ? Number.NaN : Number.parseInt(sizeValue, 10);
    const fontSize = Number.isFinite(parsedSize) && parsedSize > 0 ? parsedSize : undefined;

    let chars = 0;
    for (const textNode of findAllDeep(run, "w", "t")) {
      chars += getTextContent(textNode).length;
    }
    if (chars === 0) {
      continue;
    }

    const entry: RunMetrics = { bold, chars };
    if (fontSize !== undefined) {
      entry.fontSize = fontSize;
    }
    metrics.push(entry);
  }

  return metrics;
};

type ExtractContainerOptions = {
  container: XmlElement;
  source: DocxParagraphSource;
  startIndex: number;
};

type ExtractContainerResult = {
  paragraphs: ExtractedDocxParagraph[];
  charCount: number;
};

const extractContainer = ({
  container,
  source,
  startIndex,
}: ExtractContainerOptions): ExtractContainerResult => {
  const paragraphs: ExtractedDocxParagraph[] = [];
  let charCount = 0;

  for (const [offset, paragraph] of findAllDeep(container, "w", "p").entries()) {
    const text = collectText(paragraph);
    const entry: ExtractedDocxParagraph = {
      index: startIndex + offset,
      text,
      source,
    };
    const { style, alignment } = readParagraphProperties(paragraph);
    if (style !== undefined) {
      entry.style = style;
    }
    if (alignment !== undefined) {
      entry.alignment = alignment;
    }

    const runs = readRunMetrics(paragraph);
    if (runs.length > 0) {
      const totalChars = runs.reduce((sum, run) => sum + run.chars, 0);
      const boldChars = runs.reduce((sum, run) => sum + (run.bold ? run.chars : 0), 0);
      if (boldChars > totalChars / 2) {
        entry.bold = true;
      }
      const firstFontSize = runs.find((run) => run.fontSize !== undefined)?.fontSize;
      if (firstFontSize !== undefined) {
        entry.fontSize = firstFontSize;
      }
    }

    paragraphs.push(entry);
    charCount += text.length;
  }

  return { paragraphs, charCount };
};

type ExtractPartsOptions = {
  archive: DocxArchive;
  source: "header" | "footer";
  rootName: "hdr" | "ftr";
  startIndex: number;
  /** Part paths to read, in extraction order — see {@link resolveReferencedHeaderFooterParts}. */
  paths: readonly string[];
};

const extractParts = async ({
  archive,
  source,
  rootName,
  startIndex,
  paths,
}: ExtractPartsOptions): Promise<ExtractContainerResult> => {
  const paragraphs: ExtractedDocxParagraph[] = [];
  let charCount = 0;
  let nextIndex = startIndex;

  for (const path of paths) {
    // oxlint-disable-next-line no-await-in-loop -- part order defines stable paragraph indices
    const xml = await archive.readEntryString(path);
    if (xml === null) {
      continue;
    }
    const root = parseXml(xml);
    const container = findDeep(root, "w", rootName);
    if (!container) {
      continue;
    }
    const result = extractContainer({
      container,
      source,
      startIndex: nextIndex,
    });
    paragraphs.push(...result.paragraphs);
    charCount += result.charCount;
    nextIndex += result.paragraphs.length;
  }

  return { paragraphs, charCount };
};

/** A `word/_rels/document.xml.rels` `Target` is relative to `word/`; resolve it to a full archive-entry path. */
const resolveWordPartPath = (target: string): string =>
  target.startsWith("/") ? target.slice(1) : `word/${target}`;

type ReferencedHeaderFooterParts = {
  headers: string[];
  footers: string[];
};

/**
 * Resolve the header/footer parts actually wired into the document via
 * `word/_rels/document.xml.rels` + each section's `w:headerReference` /
 * `w:footerReference`, instead of extracting every `word/header*.xml` /
 * `word/footer*.xml` entry by filename. A DOCX can carry an orphaned
 * header/footer part (stale, or planted by an attacker) that no section
 * references — reading it unconditionally would surface prompt-injection or
 * stale content that Word itself never renders.
 */
const resolveReferencedHeaderFooterParts = async (
  archive: DocxArchive,
  documentRoot: XmlElement,
): Promise<ReferencedHeaderFooterParts> => {
  const relsXml = await archive.readEntryString(DOCUMENT_RELS_PATH);
  if (relsXml === null) {
    return { headers: [], footers: [] };
  }
  const relationships = parseRelationships(relsXml);

  const headerRIds = new Set<string>();
  const footerRIds = new Set<string>();
  for (const sectPr of findAllDeep(documentRoot, "w", "sectPr")) {
    for (const ref of findAllDeep(sectPr, "w", "headerReference")) {
      const rId = getAttribute(ref, "r", "id");
      if (rId !== null) {
        headerRIds.add(rId);
      }
    }
    for (const ref of findAllDeep(sectPr, "w", "footerReference")) {
      const rId = getAttribute(ref, "r", "id");
      if (rId !== null) {
        footerRIds.add(rId);
      }
    }
  }

  const resolvePaths = (rIds: Set<string>, relationshipType: string): string[] => {
    const paths = new Set<string>();
    for (const rId of rIds) {
      const relationship = relationships.get(rId);
      if (
        !relationship ||
        relationship.type !== relationshipType ||
        relationship.targetMode === "External"
      ) {
        continue;
      }
      paths.add(resolveWordPartPath(relationship.target));
    }
    return [...paths].toSorted();
  };

  return {
    headers: resolvePaths(headerRIds, RELATIONSHIP_TYPES.header),
    footers: resolvePaths(footerRIds, RELATIONSHIP_TYPES.footer),
  };
};

const createEmptyResult = (): ExtractedDocxText => ({
  paragraphs: [],
  charCount: 0,
  view: "accepted",
});

/** Extract paragraph text and formatting metadata from a DOCX archive. */
export const extractDocxText = async (
  bytes: ArrayBuffer | Uint8Array,
): Promise<ExtractedDocxText> => {
  const archive = await loadDocxArchive(bytes);
  const documentXml = await archive.readEntryString("word/document.xml");
  if (documentXml === null) {
    return createEmptyResult();
  }

  const root = parseXml(documentXml);
  const body = findDeep(root, "w", "body");
  if (!body) {
    return createEmptyResult();
  }

  const referencedParts = await resolveReferencedHeaderFooterParts(archive, root);

  const headers = await extractParts({
    archive,
    source: "header",
    rootName: "hdr",
    startIndex: 0,
    paths: referencedParts.headers,
  });
  const bodyResult = extractContainer({
    container: body,
    source: "body",
    startIndex: headers.paragraphs.length,
  });
  const footers = await extractParts({
    archive,
    source: "footer",
    rootName: "ftr",
    startIndex: headers.paragraphs.length + bodyResult.paragraphs.length,
    paths: referencedParts.footers,
  });

  return {
    paragraphs: [...headers.paragraphs, ...bodyResult.paragraphs, ...footers.paragraphs],
    charCount: headers.charCount + bodyResult.charCount + footers.charCount,
    view: "accepted",
  };
};
