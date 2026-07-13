import type { DocxArchive } from "./boundedArchive";
import { loadDocxArchive } from "./boundedArchive";
import {
  findAllDeep,
  findChild,
  findDeep,
  getAttributeAnyPrefix,
  getLocalName,
  getTextContent,
  parseXml,
  type XmlElement,
} from "../xmlParser";

const HEADER_FOOTER_PATH = /^word\/(?:header|footer)\d+\.xml$/u;

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
};

const extractParts = async ({
  archive,
  source,
  rootName,
  startIndex,
}: ExtractPartsOptions): Promise<ExtractContainerResult> => {
  const paragraphs: ExtractedDocxParagraph[] = [];
  let charCount = 0;
  let nextIndex = startIndex;
  const prefix = `word/${source}`;
  const paths = archive.entries
    .filter((path) => HEADER_FOOTER_PATH.test(path) && path.startsWith(prefix))
    .toSorted();

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

  const headers = await extractParts({
    archive,
    source: "header",
    rootName: "hdr",
    startIndex: 0,
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
  });

  return {
    paragraphs: [...headers.paragraphs, ...bodyResult.paragraphs, ...footers.paragraphs],
    charCount: headers.charCount + bodyResult.charCount + footers.charCount,
    view: "accepted",
  };
};
