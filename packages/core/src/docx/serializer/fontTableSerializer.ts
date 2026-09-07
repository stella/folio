import type { FontInfo, FontTable } from "../../types/document";
import { serializePartElement } from "./partNamespaces";
import { escapeXml } from "./xmlUtils";

export const serializeFontTableXml = (fontTable: FontTable): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  serializePartElement({
    partPath: "word/fontTable.xml",
    rootName: "w:fonts",
    baselinePrefixes: ["w"],
    sourceBindings: undefined,
    body: fontTable.fonts.map(serializeFont).join(""),
  });

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
    const attrs: string[] = [];
    for (const [key, value] of Object.entries(font.sig)) {
      if (value === undefined) {
        continue;
      }
      attrs.push(`w:${key}="${escapeXml(value)}"`);
    }
    if (attrs.length > 0) {
      parts.push(`<w:sig ${attrs.join(" ")}/>`);
    }
  }
  return `<w:font w:name="${escapeXml(font.name)}">${parts.join("")}</w:font>`;
};
