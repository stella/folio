import type { DocumentSettings } from "../../types/document";
import { escapeXml, intAttr } from "./xmlUtils";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const serializeSettingsXml = (settings: DocumentSettings): string => {
  const parts = [`<w:defaultTabStop w:val="${intAttr(settings.defaultTabStop)}"/>`];
  if (settings.evenAndOddHeaders) {
    parts.push("<w:evenAndOddHeaders/>");
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W_NS}">${parts.join("")}</w:settings>`;
};
