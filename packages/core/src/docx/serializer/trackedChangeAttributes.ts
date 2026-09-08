import { normalizeRevisionId } from "@stll/docx-core/model";

import type { PropertyChangeInfo, TrackedChangeInfo } from "../../types/document";
import { DATE_UTC_ATTRIBUTE } from "../trackedChangeInfo";
import { escapeXml } from "./xmlUtils";

type SerializableTrackedChangeInfo = TrackedChangeInfo | PropertyChangeInfo;

/** Normalized, unescaped attributes in schema order for an XML element or string serializer. */
export const trackedChangeAttributeEntries = (
  info: SerializableTrackedChangeInfo,
): readonly (readonly [string, string])[] => {
  const author = typeof info.author === "string" ? info.author.trim() : "";
  const date = typeof info.date === "string" ? info.date.trim() : "";
  const rawUtcDate: unknown = info.utcDate;
  const utcDate =
    typeof rawUtcDate === "object" &&
    rawUtcDate !== null &&
    "value" in rawUtcDate &&
    typeof rawUtcDate.value === "string"
      ? rawUtcDate.value.trim()
      : "";
  const rsid = "rsid" in info && typeof info.rsid === "string" ? info.rsid.trim() : "";
  const entries: [string, string][] = [
    ["w:id", String(normalizeRevisionId(info.id))],
    ["w:author", author.length > 0 ? author : "Unknown"],
  ];
  if (date.length > 0) {
    entries.push(["w:date", date]);
  }
  if (utcDate.length > 0) {
    // The model may have originated in another part whose lexical prefix is
    // no longer in scope. Rebuild the known namespace under our declared name.
    entries.push([DATE_UTC_ATTRIBUTE, utcDate]);
  }
  if (rsid.length > 0) {
    entries.push(["w:rsid", rsid]);
  }
  return entries;
};

export const trackedChangeAttributeRecord = (
  info: SerializableTrackedChangeInfo,
): Record<string, string> => Object.fromEntries(trackedChangeAttributeEntries(info));

export const serializeTrackedChangeAttributes = (info: SerializableTrackedChangeInfo): string =>
  trackedChangeAttributeEntries(info)
    .map(([name, value]) => `${name}="${escapeXml(value)}"`)
    .join(" ");
