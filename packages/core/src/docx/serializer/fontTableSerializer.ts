import type { FontInfo, FontTable } from "../../types/document";
import { escapeXml } from "./xmlUtils";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const serializeFontTableXml = (fontTable: FontTable): string => {
  const fonts = fontTable.fonts.map(serializeFont).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="${W_NS}">${fonts}</w:fonts>`;
};

const serializeFont = (font: FontInfo): string => {
  const parts: string[] = [];
  if (font.altName) {
    parts.push(`<w:altName w:val="${escapeXml(font.altName)}"/>`);
  }
  if (font.panose1) {
    parts.push(`<w:panose1 w:val="${escapeXml(font.panose1)}"/>`);
  }
  if (font.charset) {
    parts.push(`<w:charset w:val="${escapeXml(font.charset)}"/>`);
  }
  if (font.family) {
    parts.push(`<w:family w:val="${font.family}"/>`);
  }
  if (font.pitch) {
    parts.push(`<w:pitch w:val="${font.pitch}"/>`);
  }
  if (font.sig) {
    const attrs = Object.entries(font.sig)
      .map(([key, value]) => `w:${key}="${escapeXml(value)}"`)
      .join(" ");
    if (attrs.length > 0) {
      parts.push(`<w:sig ${attrs}/>`);
    }
  }
  return `<w:font w:name="${escapeXml(font.name)}">${parts.join("")}</w:font>`;
};
