import type JSZip from "jszip";
import { OOXML_NS } from "@stll/docx-utils";
import { panic } from "better-result";

import type { Document } from "../types/document";
import type { RemovedSectionReference } from "../internal/sectionEndpointResolution";
import { parseRelationships, RELATIONSHIP_TYPES, resolveRelativePath } from "./relsParser";
import { isUnsafePackagePath } from "./packageParts";
import {
  cloneElement,
  elementToXml,
  getAttribute,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  type XmlElement,
} from "./xmlParser";

const DOCUMENT_RELS_PATH = "word/_rels/document.xml.rels";

const withoutChildren = (xml: string, remove: (child: XmlElement) => boolean): string => {
  const root = parseXmlDocument(xml);
  if (!root) return panic("Cannot update malformed package metadata");
  return elementToXml(cloneElement(root, {elements: root.elements?.filter((child) => !remove(child))}));
};

type RemoveResolvedHeaderFooterPartsOptions = {
  document: Document;
  zip: JSZip;
  removedReferences: readonly RemovedSectionReference[];
  compressionLevel: number;
};

/** Remove only parts whose last selection was explicitly resolved out of the document. */
export const removeResolvedHeaderFooterParts = async ({
  document,
  zip,
  removedReferences,
  compressionLevel,
}: RemoveResolvedHeaderFooterPartsOptions): Promise<void> => {
  if (removedReferences.length === 0) return;
  const relsFile = zip.file(DOCUMENT_RELS_PATH);
  if (!relsFile) return panic("A resolved header/footer part has no document relationships");
  const relsXml = await relsFile.async("text");
  const relationships = parseRelationships(relsXml);
  const removedIds = new Set<string>();
  const candidates = new Set<string>();
  for (const {part, relationshipId} of removedReferences) {
    const parts = part === "header" ? document.package.headers : document.package.footers;
    if (parts?.has(relationshipId)) continue;
    const relationship = relationships.get(relationshipId);
    if (!relationship) continue;
    if (relationship.type !== RELATIONSHIP_TYPES[part] || relationship.targetMode === "External") {
      return panic("Resolved header/footer relationship has an unexpected part type");
    }
    const path = resolveRelativePath(DOCUMENT_RELS_PATH, relationship.target);
    if (isUnsafePackagePath(path)) return panic("Resolved header/footer has an unsafe package path");
    removedIds.add(relationshipId);
    candidates.add(path);
  }
  if (removedIds.size === 0) return;
  for (const relationship of relationships.values()) {
    if (removedIds.has(relationship.id) || relationship.targetMode === "External") continue;
    candidates.delete(resolveRelativePath(DOCUMENT_RELS_PATH, relationship.target));
  }
  // Another package part may share the retired story part through its own relationship.
  for (const file of Object.values(zip.files)) {
    if (file.dir || !file.name.endsWith(".rels") || file.name === DOCUMENT_RELS_PATH) continue;
    const otherRelationships = parseRelationships(await file.async("text"));
    for (const relationship of otherRelationships.values()) {
      if (relationship.targetMode !== "External") {
        candidates.delete(resolveRelativePath(file.name, relationship.target));
      }
    }
  }
  const retiredMedia = new Set<string>();
  for (const path of candidates) {
    const slash = path.lastIndexOf("/");
    const relsPath = `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
    const partRels = zip.file(relsPath);
    if (!partRels) continue;
    for (const relationship of parseRelationships(await partRels.async("text")).values()) {
      if (relationship.type !== RELATIONSHIP_TYPES.image || relationship.targetMode === "External") continue;
      const mediaPath = resolveRelativePath(relsPath, relationship.target);
      if (!isUnsafePackagePath(mediaPath) && mediaPath.startsWith("word/media/")) retiredMedia.add(mediaPath);
    }
  }
  const compressionOptions = {level: compressionLevel};
  zip.file(DOCUMENT_RELS_PATH, withoutChildren(relsXml, (child) =>
    getNamespaceUri(child) === OOXML_NS.pr && getLocalName(child.name) === "Relationship" &&
    removedIds.has(getAttribute(child, null, "Id") ?? "")), {compression: "DEFLATE", compressionOptions});
  if (document.package.relationships) {
    document.package.relationships = new Map(document.package.relationships);
    for (const id of removedIds) document.package.relationships.delete(id);
  }
  for (const path of candidates) {
    zip.remove(path);
    const slash = path.lastIndexOf("/");
    zip.remove(`${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`);
  }
  for (const file of Object.values(zip.files)) {
    if (file.dir || !file.name.endsWith(".rels")) continue;
    for (const relationship of parseRelationships(await file.async("text")).values()) {
      if (relationship.targetMode !== "External") retiredMedia.delete(resolveRelativePath(file.name, relationship.target));
    }
  }
  for (const path of retiredMedia) {
    zip.remove(path);
    candidates.add(path);
  }
  const contentTypes = zip.file("[Content_Types].xml");
  if (!contentTypes) return panic("The package has no content types");
  const contentTypesXml = await contentTypes.async("text");
  zip.file("[Content_Types].xml", withoutChildren(contentTypesXml, (child) =>
    getNamespaceUri(child) === OOXML_NS.ct && getLocalName(child.name) === "Override" &&
    candidates.has((getAttribute(child, null, "PartName") ?? "").replace(/^\//u, ""))),
  {compression: "DEFLATE", compressionOptions});
};
