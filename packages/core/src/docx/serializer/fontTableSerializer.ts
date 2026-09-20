import type { EmbeddedFontRef, FontInfo, FontTable } from "../../types/document";

import { serializePreservedAttributes } from "../attributeRemainder";
import { serializeWithPreservedChildren } from "../containerChildren";
import { serializePartElement } from "./partNamespaces";
import { escapeXmlAttribute } from "@stll/docx-core";

export const serializeFontTableXml = (fontTable: FontTable): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  serializePartElement({
    partPath: "word/fontTable.xml",
    rootName: "w:fonts",
    rootAttributes: serializePreservedAttributes([], fontTable.preservedAttributes).join(" "),
    baselinePrefixes: ["w"],
    sourceBindings: undefined,
    body: serializeWithPreservedChildren(fontTable.fonts.map(serializeFont), fontTable.preserved),
  });

/**
 * One `w:font`, with its children in the order `CT_Font` declares them.
 *
 * The order is what makes the sink's index mean something: a capture recorded
 * after the third field the reader kept goes back after the third field this
 * writes, so unmodelled markup keeps the neighbours it was authored between.
 */
const serializeFont = (font: FontInfo): string => {
  const parts: string[] = [];
  if (font.altName) {
    parts.push(`<w:altName w:val="${escapeXmlAttribute(font.altName)}"/>`);
  }
  if (font.panose1) {
    parts.push(`<w:panose1 w:val="${escapeXmlAttribute(font.panose1)}"/>`);
  }
  if (font.charset) {
    const attributes: string[] = [];
    if (font.charset.val !== undefined) {
      attributes.push(`w:val="${escapeXmlAttribute(font.charset.val)}"`);
    }
    if (font.charset.characterSet !== undefined) {
      attributes.push(`w:characterSet="${escapeXmlAttribute(font.charset.characterSet)}"`);
    }
    // A bare `<w:charset/>` is the default code page, which is not the same as
    // no `w:charset` at all, so the element is written whether or not it
    // states anything.
    parts.push(attributes.length === 0 ? "<w:charset/>" : `<w:charset ${attributes.join(" ")}/>`);
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
      attrs.push(`w:${key}="${escapeXmlAttribute(value)}"`);
    }
    if (attrs.length > 0) {
      parts.push(`<w:sig ${attrs.join(" ")}/>`);
    }
  }
  pushEmbed(parts, "embedRegular", font.embedRegular);
  pushEmbed(parts, "embedBold", font.embedBold);
  pushEmbed(parts, "embedItalic", font.embedItalic);
  pushEmbed(parts, "embedBoldItalic", font.embedBoldItalic);

  const attributes = serializePreservedAttributes(
    [`w:name="${escapeXmlAttribute(font.name)}"`],
    font.preservedAttributes,
  );
  const body = serializeWithPreservedChildren(parts, font.preserved);
  return `<w:font ${attributes.join(" ")}>${body}</w:font>`;
};

const pushEmbed = (parts: string[], element: string, embed: EmbeddedFontRef | undefined): void => {
  if (embed === undefined) {
    return;
  }
  const attributes = [`r:id="${escapeXmlAttribute(embed.id)}"`];
  if (embed.fontKey !== undefined) {
    attributes.push(`w:fontKey="${escapeXmlAttribute(embed.fontKey)}"`);
  }
  if (embed.subsetted !== undefined) {
    attributes.push(`w:subsetted="${embed.subsetted ? "1" : "0"}"`);
  }
  parts.push(`<w:${element} ${attributes.join(" ")}/>`);
};
