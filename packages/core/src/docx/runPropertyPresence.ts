import { EMPTY_PROPERTY_SETS } from "@stll/docx-core/ops";
import type { Run, TextFormatting, PreservedAttribute } from "../types/document";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import { getAttributeByNamespaceUri, type XmlElement } from "./xmlParser";

export const EMPTY_RUN_PROPERTIES_ATTRIBUTE = "emptyRunProperties";
const PREVIOUS_PROPERTIES_ATTRIBUTE = "previousProperties";
const HISTORY_NAMESPACES: ReadonlySet<string> = new Set([OOXML_NAMESPACES.folio.uri]);

export const retainsEmptyRunProperties = (element: XmlElement): boolean =>
  getAttributeByNamespaceUri(element, HISTORY_NAMESPACES, EMPTY_RUN_PROPERTIES_ATTRIBUTE) ===
  EMPTY_PROPERTY_SETS.KEEP;

export const omitsPreviousRunProperties = (element: XmlElement): boolean =>
  getAttributeByNamespaceUri(element, HISTORY_NAMESPACES, PREVIOUS_PROPERTIES_ATTRIBUTE) ===
  EMPTY_PROPERTY_SETS.OMIT;

export const serializePreviousRunProperties = (properties: TextFormatting | undefined): string =>
  properties === undefined
    ? ` folio:${PREVIOUS_PROPERTIES_ATTRIBUTE}="${EMPTY_PROPERTY_SETS.OMIT}"`
    : "";

export const serializeEmptyRunProperties = (run: Run): string[] =>
  run.propertyChanges?.length &&
  run.formatting !== undefined &&
  Object.keys(run.formatting).length === 0
    ? [`folio:${EMPTY_RUN_PROPERTIES_ATTRIBUTE}="${EMPTY_PROPERTY_SETS.KEEP}"`]
    : [];

export const isEmptyRunPropertiesAttribute = ({
  namespace,
  name,
}: Pick<PreservedAttribute, "namespace" | "name">): boolean =>
  namespace === OOXML_NAMESPACES.folio.uri && name === EMPTY_RUN_PROPERTIES_ATTRIBUTE;
