import type { Mark, Node as PMNode } from "prosemirror-model";
import { expectPreservedXmlAttrs, expectTrackedChangeMarkAttrs } from "./attrs";
import { captureVerbatimXml } from "../docx/verbatimCapture";
import {
  cloneElement,
  getChildElements,
  getLocalName,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "../docx/xmlParser";

const restoreDeletedText = (node: XmlElement): XmlElement => {
  if (node.type !== "element") return node;
  const name = getLocalName(node.name);
  let target: string | undefined;
  if (name === "delText") target = "t";
  if (name === "delInstrText") target = "instrText";
  const renamed =
    target !== undefined && WORDPROCESSINGML_NAMESPACE_URIS.has(node.namespaceUri ?? "")
      ? node.name?.replace(/[^:]+$/u, target)
      : node.name;
  const elements = node.elements?.map(restoreDeletedText);
  if (
    renamed === node.name &&
    elements?.every((child, index) => child === node.elements?.[index]) !== false
  )
    return node;
  return cloneElement(node, {
    ...(renamed !== node.name ? { name: renamed } : {}),
    ...(elements ? { elements } : {}),
  });
};

const hasPendingDeletion = (marks: readonly Mark[]) =>
  marks.some((mark) => {
    if (mark.type.name === "deletion") return true;
    if (mark.type.name !== "insertion") return false;
    return (
      expectTrackedChangeMarkAttrs(mark)._docxRevisionAncestors?.some(
        (ancestor) => ancestor.type === "deletion" || ancestor.type === "moveFrom",
      ) ?? false
    );
  });

/** A retained opaque deleted result becomes ordinary text when its last deletion resolves. */
export const resolvedPreservedXmlAttrs = (node: PMNode, marks: readonly Mark[]) => {
  if (
    node.type.name !== "preservedXml" ||
    !hasPendingDeletion(node.marks) ||
    hasPendingDeletion(marks)
  )
    return node.attrs;
  const attrs = expectPreservedXmlAttrs(node);
  const xml = getChildElements(parseXml(attrs.xml, OOXML_NAMESPACE_SCOPE))
    .map((element) => captureVerbatimXml(restoreDeletedText(element)))
    .join("");
  return xml === attrs.xml ? node.attrs : { ...node.attrs, xml };
};
