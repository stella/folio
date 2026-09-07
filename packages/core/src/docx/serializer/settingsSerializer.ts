import type { DocumentSettings } from "../../types/document";
import { serializePartElement } from "./partNamespaces";
import { escapeXml, intAttr } from "./xmlUtils";

export const serializeSettingsXml = (settings: DocumentSettings): string => {
  const parts = [`<w:defaultTabStop w:val="${intAttr(settings.defaultTabStop)}"/>`];
  if (settings.evenAndOddHeaders) {
    parts.push("<w:evenAndOddHeaders/>");
  }
  if (settings.updateFields) {
    parts.push('<w:updateFields w:val="true"/>');
  }
  if (settings.themeFontLang) {
    const attrs: string[] = [];
    if (settings.themeFontLang.eastAsia) {
      attrs.push(`w:eastAsia="${escapeXml(settings.themeFontLang.eastAsia)}"`);
    }
    if (settings.themeFontLang.bidi) {
      attrs.push(`w:bidi="${escapeXml(settings.themeFontLang.bidi)}"`);
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
