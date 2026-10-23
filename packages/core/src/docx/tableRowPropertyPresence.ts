import { EMPTY_PROPERTY_SETS } from "@stll/docx-core/ops";

import type { PreservedAttribute, TableRow } from "../types/document";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import { getAttributeByNamespaceUri, type XmlElement } from "./xmlParser";

export const EMPTY_ROW_PROPERTIES_ATTRIBUTE = "emptyRowProperties";
const HISTORY_NAMESPACES: ReadonlySet<string> = new Set([
  OOXML_NAMESPACES.folio.uri,
]);

export const retainsEmptyRowProperties = (element: XmlElement): boolean =>
  getAttributeByNamespaceUri(
    element,
    HISTORY_NAMESPACES,
    EMPTY_ROW_PROPERTIES_ATTRIBUTE,
  ) === EMPTY_PROPERTY_SETS.KEEP;

export const serializeEmptyRowProperties = (row: TableRow): string[] => {
  if (
    row.formatting === undefined ||
    (row.structuralChange === undefined && !row.propertyChanges?.length)
  )
    return [];
  const { sourceXml: _source, ...properties } = row.formatting;
  return Object.keys(properties).length === 0
    ? [`folio:${EMPTY_ROW_PROPERTIES_ATTRIBUTE}="${EMPTY_PROPERTY_SETS.KEEP}"`]
    : [];
};

export const isEmptyRowPropertiesAttribute = ({
  namespace,
  name,
}: Pick<PreservedAttribute, "namespace" | "name">): boolean =>
  namespace === OOXML_NAMESPACES.folio.uri &&
  name === EMPTY_ROW_PROPERTIES_ATTRIBUTE;
