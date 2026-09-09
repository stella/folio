import { normalizeRevisionId } from "@stll/docx-core/model";
import { panic } from "better-result";

import type {
  PropertyChangeInfo,
  RunPropertyChange,
  TrackedChangeInfo,
} from "../../types/document";
import { DATE_UTC_ATTRIBUTE } from "../trackedChangeInfo";
import { escapeXml } from "./xmlUtils";

type SerializableTrackedChangeInfo = TrackedChangeInfo | PropertyChangeInfo;

/** Enforces the singular `w:rPrChange` child in a run-property container. */
export const getSingularRunPropertyChange = (
  propertyChanges: readonly RunPropertyChange[] | undefined,
): RunPropertyChange | undefined => {
  const propertyChangeCount = propertyChanges?.length ?? 0;
  if (propertyChangeCount > 1) {
    panic("A run-property container cannot serialize more than one w:rPrChange", {
      elementName: "w:rPrChange",
      propertyChangeCount,
    });
  }
  return propertyChanges?.at(0);
};

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
