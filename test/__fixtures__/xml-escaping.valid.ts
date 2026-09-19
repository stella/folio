/** A writer that escapes through the owner. */

import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";

export const serializeAlias = (alias: string, text: string): string =>
  `<w:alias w:val="${escapeXmlAttribute(alias)}">${escapeXmlText(text)}</w:alias>`;

/** Reading an entity is not escaping one. */
export const mentionsEntity = (xml: string): boolean => xml.includes("&amp;");
