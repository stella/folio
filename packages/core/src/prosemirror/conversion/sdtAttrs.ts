import type { SdtProperties } from "../../types/document";
import type { SdtAttrs } from "../schema/nodes";

export const sdtAttrsFromProperties = (properties: SdtProperties): SdtAttrs => ({
  sdtType: properties.sdtType,
  ...(properties.alias !== undefined ? { alias: properties.alias } : {}),
  ...(properties.tag !== undefined ? { tag: properties.tag } : {}),
  ...(properties.id !== undefined ? { id: properties.id } : {}),
  ...(properties.lock !== undefined ? { lock: properties.lock } : {}),
  ...(properties.placeholder !== undefined ? { placeholder: properties.placeholder } : {}),
  showingPlaceholder: properties.showingPlaceholder ?? false,
  ...(properties.dateFormat !== undefined ? { dateFormat: properties.dateFormat } : {}),
  ...(properties.dateValueISO !== undefined ? { dateValueISO: properties.dateValueISO } : {}),
  ...(properties.listItems !== undefined
    ? { listItems: JSON.stringify(properties.listItems) }
    : {}),
  ...(properties.dropdownLastValue !== undefined
    ? { dropdownLastValue: properties.dropdownLastValue }
    : {}),
  ...(properties.checked !== undefined ? { checked: properties.checked } : {}),
  ...(properties.rawPropertiesXml !== undefined
    ? { rawPropertiesXml: properties.rawPropertiesXml }
    : {}),
  ...(properties.rawEndPropertiesXml !== undefined
    ? { rawEndPropertiesXml: properties.rawEndPropertiesXml }
    : {}),
});

export const sdtPropertiesFromAttrs = (attrs: SdtAttrs): SdtProperties => {
  const properties: SdtProperties = { sdtType: attrs.sdtType };
  if (attrs.alias) {
    properties.alias = attrs.alias;
  }
  if (attrs.tag) {
    properties.tag = attrs.tag;
  }
  if (typeof attrs.id === "number") {
    properties.id = attrs.id;
  }
  if (attrs.lock) {
    properties.lock = attrs.lock;
  }
  if (attrs.placeholder) {
    properties.placeholder = attrs.placeholder;
  }
  if (attrs.showingPlaceholder !== undefined) {
    properties.showingPlaceholder = attrs.showingPlaceholder;
  }
  if (attrs.dateFormat) {
    properties.dateFormat = attrs.dateFormat;
  }
  if (attrs.dateValueISO) {
    properties.dateValueISO = attrs.dateValueISO;
  }
  if (attrs.listItems) {
    properties.listItems = parseSdtListItems(attrs.listItems);
  }
  if (typeof attrs.dropdownLastValue === "string") {
    properties.dropdownLastValue = attrs.dropdownLastValue;
  }
  if (typeof attrs.checked === "boolean") {
    properties.checked = attrs.checked;
  }
  if (attrs.rawPropertiesXml) {
    properties.rawPropertiesXml = attrs.rawPropertiesXml;
  }
  if (attrs.rawEndPropertiesXml) {
    properties.rawEndPropertiesXml = attrs.rawEndPropertiesXml;
  }
  return properties;
};

export const sdtPropertiesMatchAttrs = (properties: SdtProperties, attrs: SdtAttrs): boolean =>
  JSON.stringify(sdtAttrsFromProperties(properties)) ===
  JSON.stringify(sdtAttrsFromProperties(sdtPropertiesFromAttrs(attrs)));

const parseSdtListItems = (rawItems: string): NonNullable<SdtProperties["listItems"]> => {
  const parsed = JSON.parse(rawItems) as unknown;
  if (!Array.isArray(parsed)) {
    throw new TypeError("Invalid ProseMirror sdt attrs: listItems is not an array");
  }

  return parsed.map((item): NonNullable<SdtProperties["listItems"]>[number] => {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      throw new TypeError("Invalid ProseMirror sdt attrs: listItems contains an invalid item");
    }
    if (
      !("displayText" in item) ||
      typeof item.displayText !== "string" ||
      !("value" in item) ||
      typeof item.value !== "string"
    ) {
      throw new TypeError("Invalid ProseMirror sdt attrs: listItems contains an invalid item");
    }
    return { displayText: item.displayText, value: item.value };
  });
};
