import { TaggedError } from "better-result";

import type { HeaderReference, SectionProperties, SectionPropertyChange } from "../types/document";
import {
  findChildrenByNamespaceUri,
  getAttributeByNamespaceUri,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";
import { canonicalJson } from "../utils/canonicalJson";
import { escapeXmlAttribute } from "@stll/docx-core";

const SECTION_REFERENCE_HISTORY_NAMESPACE = "urn:stella:folio:section-reference-history:1";
const HISTORY_NAMESPACES: ReadonlySet<string> = new Set([SECTION_REFERENCE_HISTORY_NAMESPACE]);

type PreviousReferences = NonNullable<SectionPropertyChange["previousReferences"]>;

export class InvalidSectionReferenceHistoryError extends TaggedError(
  "InvalidSectionReferenceHistoryError",
)<{
  message: string;
}> {}

const invalidHistory = (): never => {
  throw new InvalidSectionReferenceHistoryError({
    message: "Invalid section reference revision history.",
  });
};

export const parseSectionReferenceHistory = (
  change: XmlElement,
): PreviousReferences | undefined => {
  const histories = findChildrenByNamespaceUri(change, HISTORY_NAMESPACES, "previousReferences");
  const history = histories.at(0);
  if (!history) return undefined;
  if (histories.length !== 1) return invalidHistory();
  const headers: HeaderReference[] = [];
  const footers: HeaderReference[] = [];
  for (const child of getChildElements(history)) {
    if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")) return invalidHistory();
    const name = getLocalName(child.name);
    if (name !== "headerReference" && name !== "footerReference") return invalidHistory();
    const type = getAttributeByNamespaceUri(child, WORDPROCESSINGML_NAMESPACE_URIS, "type");
    const rId = getAttributeByNamespaceUri(child, OFFICE_RELATIONSHIP_NAMESPACE_URIS, "id");
    if ((type !== "default" && type !== "first" && type !== "even") || !rId || rId.trim() !== rId) {
      return invalidHistory();
    }
    const references = name === "headerReference" ? headers : footers;
    if (references.some((reference) => reference.type === type)) return invalidHistory();
    references.push({ type, rId });
  }
  return {
    ...(headers.length > 0 && { headerReferences: headers }),
    ...(footers.length > 0 && { footerReferences: footers }),
  };
};

export const serializeSectionReferenceHistory = (
  references: PreviousReferences | undefined,
): string => {
  if (references === undefined) return "";
  const serializeReferences = (kind: "header" | "footer", values: readonly HeaderReference[]) => {
    if (new Set(values.map(({ type }) => type)).size !== values.length) return invalidHistory();
    return values
      .map(({ type, rId }) => {
        if (!rId || rId.trim() !== rId) return invalidHistory();
        return `<w:${kind}Reference w:type="${type}" r:id="${escapeXmlAttribute(rId)}"/>`;
      })
      .join("");
  };
  const content =
    serializeReferences("header", references.headerReferences ?? []) +
    serializeReferences("footer", references.footerReferences ?? []);
  return `<frh:previousReferences xmlns:frh="${SECTION_REFERENCE_HISTORY_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="frh" mc:PreserveElements="frh:previousReferences">${content}</frh:previousReferences>`;
};

export const sectionReferenceSelection = (properties: SectionProperties): PreviousReferences => ({
  ...(properties.headerReferences !== undefined && {
    headerReferences: properties.headerReferences,
  }),
  ...(properties.footerReferences !== undefined && {
    footerReferences: properties.footerReferences,
  }),
});

export const sectionReferenceHistory = ({
  previous,
  target,
}: {
  previous: SectionProperties;
  target: SectionProperties;
}): PreviousReferences | undefined => {
  const before = sectionReferenceSelection(previous);
  return canonicalJson(before) === canonicalJson(sectionReferenceSelection(target))
    ? undefined
    : before;
};
