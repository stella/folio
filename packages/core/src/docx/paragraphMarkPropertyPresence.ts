import { EMPTY_PROPERTY_SETS } from "@stll/docx-core/ops";
import type { Paragraph } from "../types/document";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import { getAttributeByNamespaceUri, type XmlElement } from "./xmlParser";

export const EMPTY_MARK_PROPERTIES_ATTRIBUTE = "emptyMarkProperties";
const REVIEW_HISTORY_NAMESPACES: ReadonlySet<string> = new Set([OOXML_NAMESPACES.folio.uri]);

/** A paragraph-owned rPr carrier cannot distinguish absent and explicitly empty formatting. */
export const serializeEmptyMarkProperties = (paragraph: Paragraph): string | undefined => {
  const properties = paragraph.formatting?.runProperties;
  const hasCarrier =
    paragraph.pPrMark !== undefined || paragraph.formatting?.runInWithNext !== undefined;
  if (!hasCarrier || properties === undefined || Object.keys(properties).length !== 0)
    return undefined;
  return `folio:${EMPTY_MARK_PROPERTIES_ATTRIBUTE}="${EMPTY_PROPERTY_SETS.KEEP}"`;
};

export const restoreEmptyMarkProperties = (paragraph: Paragraph, element: XmlElement): void => {
  if (
    (paragraph.pPrMark === undefined && paragraph.formatting?.runInWithNext === undefined) ||
    getAttributeByNamespaceUri(
      element,
      REVIEW_HISTORY_NAMESPACES,
      EMPTY_MARK_PROPERTIES_ATTRIBUTE,
    ) !== EMPTY_PROPERTY_SETS.KEEP ||
    paragraph.formatting?.runProperties !== undefined
  )
    return;
  paragraph.formatting = { ...paragraph.formatting, runProperties: {} };
};
