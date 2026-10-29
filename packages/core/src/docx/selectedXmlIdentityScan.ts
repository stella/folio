import { scanStreamingXmlElements } from "@stll/docx-core/xml";

import {
  getLocalName,
  getNamespaceUri,
  NAMESPACES,
  selectAlternateContentBranch,
  type XmlElement,
} from "./xmlParser";

const isCompatibilityElement = (element: XmlElement, name: string) =>
  getNamespaceUri(element) === NAMESPACES.mc &&
  getLocalName(element.name) === name;

/** Identity census follows the parser's first-Choice, otherwise-Fallback policy. */
export const scanSelectedStreamingXmlElements = (
  xml: string,
  visitor: Parameters<typeof scanStreamingXmlElements>[1],
):
  | {
      status: "unsupported";
    }
  | {
      status: "scanned";
    } => {
  if (!xml.includes("AlternateContent"))
    return scanStreamingXmlElements(xml, visitor);
  const alternatives = new Map<number, XmlElement>();
  const positions = new WeakMap<XmlElement, number>();
  const scanned = scanStreamingXmlElements(
    xml,
    ({ element, parent, nameEnd }) => {
      positions.set(element, nameEnd);
      if (isCompatibilityElement(element, "AlternateContent")) {
        element.elements = [];
        alternatives.set(nameEnd, element);
      }
      if (!parent || !isCompatibilityElement(parent, "AlternateContent"))
        return;
      parent.elements?.push(element);
    },
  );
  if (scanned.status === "unsupported") return scanned;
  const selectedPositions = new Map<number, number | undefined>();
  for (const [position, alternative] of alternatives) {
    const selected = selectAlternateContentBranch(alternative);
    selectedPositions.set(
      position,
      selected === undefined ? undefined : positions.get(selected),
    );
  }
  const active = new WeakMap<XmlElement, boolean>();
  return scanStreamingXmlElements(xml, (entry) => {
    const { element, parent, nameEnd } = entry;
    positions.set(element, nameEnd);
    let included = parent === undefined || active.get(parent) === true;
    if (parent && isCompatibilityElement(parent, "AlternateContent")) {
      included &&=
        nameEnd === selectedPositions.get(positions.get(parent) ?? -1);
    }
    active.set(element, included);
    if (included) visitor(entry);
  });
};
