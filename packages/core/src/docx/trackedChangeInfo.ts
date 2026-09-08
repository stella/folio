import { normalizeRevisionId } from "@stll/docx-core/model";

import type { PropertyChangeInfo, TrackedChangeInfo } from "../types/document";
import {
  findAttributeByNamespaceUri,
  getAttributeByNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

export const DATE_UTC_NAMESPACE_URI =
  "http://schemas.microsoft.com/office/word/2023/wordml/word16du";
export const DATE_UTC_ATTRIBUTE = "w16du:dateUtc";

const DATE_UTC_NAMESPACE_URIS: ReadonlySet<string> = new Set([DATE_UTC_NAMESPACE_URI]);

/** Parse the metadata shared by every WordprocessingML tracked-change element. */
export const parseTrackedChangeInfo = (node: XmlElement): TrackedChangeInfo => {
  const rawId = getAttributeByNamespaceUri(node, WORDPROCESSINGML_NAMESPACE_URIS, "id");
  const parsedId = rawId ? Number.parseInt(rawId, 10) : 0;
  const author = (
    getAttributeByNamespaceUri(node, WORDPROCESSINGML_NAMESPACE_URIS, "author") ?? ""
  ).trim();
  const date = (
    getAttributeByNamespaceUri(node, WORDPROCESSINGML_NAMESPACE_URIS, "date") ?? ""
  ).trim();
  const initials = (
    getAttributeByNamespaceUri(node, WORDPROCESSINGML_NAMESPACE_URIS, "initials") ?? ""
  ).trim();

  const info: TrackedChangeInfo = {
    // `w:id` is attacker-controlled and unbounded in the schema.
    id: normalizeRevisionId(parsedId),
    author: author.length > 0 ? author : "Unknown",
  };
  if (date.length > 0) {
    info.date = date;
  }
  if (initials.length > 0) {
    info.initials = initials;
  }
  const utcDate = findAttributeByNamespaceUri(node, DATE_UTC_NAMESPACE_URIS, "dateUtc");
  const utcDateValue = utcDate?.value.trim() ?? "";
  if (utcDate && utcDateValue.length > 0) {
    // The source prefix is scoped to the parsed part. Keep only its namespace
    // identity so rebuilt parts cannot replay an unbound or shadowed prefix.
    info.utcDate = { attribute: DATE_UTC_ATTRIBUTE, value: utcDateValue };
  }
  return info;
};

/** Parse the shared metadata plus the optional property-change revision session. */
export const parsePropertyChangeInfo = (node: XmlElement): PropertyChangeInfo => {
  const info = parseTrackedChangeInfo(node);
  const rsid = (
    getAttributeByNamespaceUri(node, WORDPROCESSINGML_NAMESPACE_URIS, "rsid") ?? ""
  ).trim();
  return rsid.length > 0 ? { ...info, rsid } : info;
};
