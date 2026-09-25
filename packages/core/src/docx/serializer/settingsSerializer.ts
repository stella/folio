import type { DocumentSettings } from "../../types/document";
import { serializePartElement } from "./partNamespaces";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute, serializeOnOffElement } from "@stll/docx-core";

/** `w:compatSetting/@w:uri` the named application compatibility settings share. */
const COMPAT_SETTING_URI = "http://schemas.microsoft.com/office/word";

/**
 * Named compatibility settings written on alongside `compatibilityMode` 15 and
 * later. Layout never gates run typography on a compatibility setting and
 * never flips `w:mirrorIndents`, so declaring both on describes how the
 * package is already laid out.
 */
const MODERN_COMPAT_SETTINGS = ["enableOpenTypeFeatures", "doNotFlipMirrorIndents"] as const;

const FIRST_MODERN_COMPATIBILITY_MODE = 15;

const compatSettingXml = (name: string, val: number): string =>
  `<w:compatSetting w:name="${name}" w:uri="${COMPAT_SETTING_URI}" w:val="${intAttr(val)}"/>`;

const serializeCompat = (compatibilityMode: number | undefined): string => {
  if (compatibilityMode === undefined) {
    return "";
  }
  const settings = [compatSettingXml("compatibilityMode", compatibilityMode)];
  if (compatibilityMode >= FIRST_MODERN_COMPATIBILITY_MODE) {
    for (const name of MODERN_COMPAT_SETTINGS) {
      settings.push(compatSettingXml(name, 1));
    }
  }
  return `<w:compat>${settings.join("")}</w:compat>`;
};

export const serializeSettingsXml = (settings: DocumentSettings): string => {
  const parts = [
    `<w:defaultTabStop w:val="${intAttr(settings.defaultTabStop)}"/>`,
    serializeOnOffElement(settings.evenAndOddHeaders, "evenAndOddHeaders"),
    serializeOnOffElement(settings.updateFields, "updateFields"),
    serializeCompat(settings.compatibilityMode),
  ];
  if (settings.themeFontLang) {
    const attrs: string[] = [];
    if (settings.themeFontLang.eastAsia) {
      attrs.push(`w:eastAsia="${escapeXmlAttribute(settings.themeFontLang.eastAsia)}"`);
    }
    if (settings.themeFontLang.bidi) {
      attrs.push(`w:bidi="${escapeXmlAttribute(settings.themeFontLang.bidi)}"`);
    }
    if (attrs.length > 0) {
      parts.push(`<w:themeFontLang ${attrs.join(" ")}/>`);
    }
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    serializePartElement({
      partPath: "word/settings.xml",
      rootName: "w:settings",
      baselinePrefixes: ["w"],
      sourceBindings: undefined,
      body: parts.join(""),
    })
  );
};
