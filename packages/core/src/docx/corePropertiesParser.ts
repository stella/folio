import type { DocxPackage } from "../types/document";
import { findChildByLocalName, getTextContent, parseXmlDocument } from "./xmlParser";

type DocumentProperties = NonNullable<DocxPackage["properties"]>;

const STRING_PROPERTIES = [
  "title",
  "subject",
  "creator",
  "keywords",
  "description",
  "lastModifiedBy",
] as const;

const parseRevision = (value: string): number | undefined => {
  if (!/^\d+$/u.test(value)) {
    return undefined;
  }
  const revision = Number(value);
  return Number.isSafeInteger(revision) ? revision : undefined;
};

const parseDate = (value: string): Date | undefined => {
  const timestamp = Date.parse(value);
  return Number.isNaN(timestamp) ? undefined : new Date(timestamp);
};

/** Parse recognized package metadata from `docProps/core.xml`. */
export const parseCoreProperties = (xml: string | null): DocumentProperties | undefined => {
  if (!xml) {
    return undefined;
  }
  const root = parseXmlDocument(xml);
  if (!root) {
    return undefined;
  }

  const properties: DocumentProperties = {};
  let recognizedPropertyCount = 0;
  for (const property of STRING_PROPERTIES) {
    const element = findChildByLocalName(root, property);
    if (!element) {
      continue;
    }
    properties[property] = getTextContent(element);
    recognizedPropertyCount++;
  }

  const revisionElement = findChildByLocalName(root, "revision");
  if (revisionElement) {
    const revision = parseRevision(getTextContent(revisionElement));
    if (revision !== undefined) {
      properties.revision = revision;
      recognizedPropertyCount++;
    }
  }

  for (const property of ["created", "modified"] as const) {
    const element = findChildByLocalName(root, property);
    if (!element) {
      continue;
    }
    const date = parseDate(getTextContent(element));
    if (date !== undefined) {
      properties[property] = date;
      recognizedPropertyCount++;
    }
  }

  return recognizedPropertyCount > 0 ? properties : undefined;
};
