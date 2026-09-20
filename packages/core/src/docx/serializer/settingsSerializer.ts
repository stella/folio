import type { DocumentSettings } from "../../types/document";
import { serializePartElement } from "./partNamespaces";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute } from "@stll/docx-core";

export const serializeSettingsXml = (settings: DocumentSettings): string => {
  const parts = [`<w:defaultTabStop w:val="${intAttr(settings.defaultTabStop)}"/>`];
  if (settings.evenAndOddHeaders) {
    parts.push("<w:evenAndOddHeaders/>");
  }
  if (settings.updateFields) {
    // `CT_OnOff` with no `w:val` is an on, the spelling the rest of the
    // package writes; `w:evenAndOddHeaders` above already uses it.
    parts.push("<w:updateFields/>");
  }
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
