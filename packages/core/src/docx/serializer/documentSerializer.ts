/**
 * Document Serializer - Serialize complete document.xml
 *
 * Converts Document objects back to valid document.xml OOXML format.
 * Combines all content (paragraphs, tables) with section properties
 * and proper namespace declarations.
 *
 * OOXML Reference:
 * - Document root: w:document
 * - Document body: w:body
 * - Section properties: w:sectPr
 */

import type {
  Document,
  DocumentBackground,
  DocumentBody,
  BlockContent,
} from "../../types/document";
import { escapeXmlAttribute } from "@stll/docx-core";
import { themeColorToken } from "@stll/docx-core/model";
import { serializeBlockSdt } from "./blockSdtSerializer";
import { serializeBookmarkMarker } from "./markupRangeAttributes";
import { serializePartElement, type OoxmlNamespacePrefix } from "./partNamespaces";
import { serializeParagraph } from "./paragraphSerializer";
import { resetAutoIdCounter } from "./runSerializer";
import { serializeSectionProperties } from "./sectionPropertiesSerializer";
import { serializeTable } from "./tableSerializer";

/**
 * Prefixes document.xml declares whether or not the body uses them.
 *
 * A canonical `<w:sdtPr>` legitimately carries `<w16sdtdh:dataHash>` /
 * `<w16cex:*>` / `<w16cid:*>` children the parser stores opaquely, and the
 * drawing prefixes host raw-replayed shapes, so the part keeps the shape Word
 * writes even when this particular save emits none of them.
 */
const DOCUMENT_BASELINE_PREFIXES = [
  "a",
  "wpc",
  "mc",
  "o",
  "pic",
  "r",
  "m",
  "v",
  "wp14",
  "wp",
  "w10",
  "w",
  "w14",
  "w15",
  "w16",
  "w16cex",
  "w16cid",
  "w16sdtdh",
  "w16se",
  "wpg",
  "wps",
] as const satisfies readonly OoxmlNamespacePrefix[];

// ============================================================================
// CONTENT SERIALIZATION
// ============================================================================

/** Serialize a single block, whether or not folio models what is in it. */
function serializeBlockContent(block: BlockContent): string {
  switch (block.type) {
    case "paragraph":
      return serializeParagraph(block);
    case "table":
      return serializeTable(block, serializeParagraph);
    case "blockSdt":
      return serializeBlockSdt(block, serializeBlockContent);
    case "preservedBlock":
      return block.xml;
    // The container declares the marker beside its blocks and the model keeps
    // it there, so the save writes it back between the same two siblings.
    case "bookmarkStart":
    case "bookmarkEnd":
      return serializeBookmarkMarker(block);
    default: {
      const unreachable: never = block;
      return unreachable;
    }
  }
}

/**
 * Serialize document body content
 */
function serializeBodyContent(content: BlockContent[]): string {
  return content.map((block) => serializeBlockContent(block)).join("");
}

/** Serialize the optional page background that precedes `w:body`. */
function serializeDocumentBackground(background: DocumentBackground | undefined): string {
  if (background === undefined) {
    return "";
  }

  const attributes: string[] = [];
  if (background.color?.auto) {
    attributes.push('w:color="auto"');
  } else if (background.color?.rgb !== undefined) {
    attributes.push(`w:color="${escapeXmlAttribute(background.color.rgb)}"`);
  }
  if (background.themeColor !== undefined) {
    attributes.push(`w:themeColor="${escapeXmlAttribute(themeColorToken(background.themeColor))}"`);
  }
  if (background.themeTint !== undefined) {
    attributes.push(`w:themeTint="${escapeXmlAttribute(background.themeTint)}"`);
  }
  if (background.themeShade !== undefined) {
    attributes.push(`w:themeShade="${escapeXmlAttribute(background.themeShade)}"`);
  }

  const attributeText = attributes.length === 0 ? "" : ` ${attributes.join(" ")}`;
  const drawing = background.drawing?.rawXml;
  return drawing === undefined
    ? `<w:background${attributeText}/>`
    : `<w:background${attributeText}>${drawing}</w:background>`;
}

// ============================================================================
// MAIN DOCUMENT SERIALIZATION
// ============================================================================

/**
 * Serialize a DocumentBody to document.xml body content
 *
 * @param body - The document body to serialize
 * @returns XML string for the body element (without body tags)
 */
export function serializeDocumentBody(body: DocumentBody): string {
  const parts: string[] = [];

  // Serialize all content blocks
  parts.push(serializeBodyContent(body.content));

  // Final section properties (at the end of body)
  if (body.finalSectionProperties) {
    parts.push(serializeSectionProperties(body.finalSectionProperties));
  }

  return parts.join("");
}

/**
 * Serialize a complete Document to valid document.xml
 *
 * @param doc - The document to serialize
 * @param sourceBindings - Root `xmlns:*` of the part being replaced, so a
 *   prefix only the source document bound keeps its URI
 * @returns Complete XML string for document.xml
 */
export function serializeDocument(
  doc: Document,
  sourceBindings?: ReadonlyMap<string, string>,
): string {
  // Reset auto-incrementing image/shape ID counter for this serialization pass
  resetAutoIdCounter();

  const body =
    serializeDocumentBackground(doc.package.document.background) +
    `<w:body>${serializeDocumentBody(doc.package.document)}</w:body>`;

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    serializePartElement({
      partPath: "word/document.xml",
      rootName: "w:document",
      baselinePrefixes: DOCUMENT_BASELINE_PREFIXES,
      sourceBindings,
      body,
    })
  );
}

/**
 * Serialize just the document body (useful for partial updates)
 *
 * @param body - The document body to serialize
 * @returns XML string for the w:body element
 */
export function serializeDocumentBodyElement(body: DocumentBody): string {
  return `<w:body>${serializeDocumentBody(body)}</w:body>`;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if document has any content
 */
export function hasDocumentContent(doc: Document): boolean {
  return doc.package.document.content.length > 0;
}

/**
 * Check if document has sections
 */
export function hasDocumentSections(doc: Document): boolean {
  return (doc.package.document.sections?.length ?? 0) > 0;
}

/**
 * Check if document has section properties
 */
export function hasSectionProperties(doc: Document): boolean {
  return doc.package.document.finalSectionProperties !== undefined;
}

/**
 * Get document content count (paragraphs + tables)
 */
export function getDocumentContentCount(doc: Document): number {
  return doc.package.document.content.length;
}

/**
 * Get paragraph count in document
 */
export function getDocumentParagraphCount(doc: Document): number {
  return doc.package.document.content.filter((b) => b.type === "paragraph").length;
}

/**
 * Get table count in document
 */
export function getDocumentTableCount(doc: Document): number {
  return doc.package.document.content.filter((b) => b.type === "table").length;
}

/**
 * Create an empty document
 */
export function createEmptyDocument(): Document {
  return {
    package: {
      document: {
        content: [],
      },
    },
  };
}

/**
 * Create a simple document with text content
 */
export function createSimpleDocument(paragraphs: { text: string; styleId?: string }[]): Document {
  return {
    package: {
      document: {
        content: paragraphs.map((p) => ({
          type: "paragraph" as const,
          ...(p.styleId ? { formatting: { styleId: p.styleId } } : {}),
          content: [
            {
              type: "run" as const,
              content: [{ type: "text" as const, text: p.text }],
            },
          ],
        })),
      },
    },
  };
}
