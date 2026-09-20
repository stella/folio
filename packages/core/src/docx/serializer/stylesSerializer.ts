import type { Style, StyleDefinitions } from "../../types/document";
import { serializePartElement } from "./partNamespaces";
import { serializeParagraphFormatting } from "./paragraphSerializer";
import { serializeTextFormatting } from "./textFormattingSerializer";
import {
  serializeTableCellFormatting,
  serializeTableFormatting,
  serializeTableRowFormatting,
} from "./tableSerializer";
import { escapeXmlAttribute, pushOnOffElement } from "@stll/docx-core";

export const serializeStylesXml = (definitions: StyleDefinitions): string => {
  const docDefaults = serializeDocumentDefaults(definitions);
  const latentStyles = serializeLatentStyles(definitions);
  const styles = definitions.styles.map(serializeStyle).join("");
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    serializePartElement({
      partPath: "word/styles.xml",
      rootName: "w:styles",
      baselinePrefixes: ["w"],
      sourceBindings: undefined,
      body: `${docDefaults}${latentStyles}${styles}`,
    })
  );
};

const serializeDocumentDefaults = (definitions: StyleDefinitions): string => {
  const rPr = serializeTextFormatting(definitions.docDefaults?.rPr) || "<w:rPr/>";
  const pPr = serializeParagraphFormatting(definitions.docDefaults?.pPr) || "<w:pPr/>";
  return `<w:docDefaults><w:rPrDefault>${rPr}</w:rPrDefault><w:pPrDefault>${pPr}</w:pPrDefault></w:docDefaults>`;
};

const serializeLatentStyles = (definitions: StyleDefinitions): string => {
  const latent = definitions.latentStyles;
  if (!latent) {
    return "";
  }
  const attrs: string[] = [];
  pushBooleanAttr(attrs, "w:defLockedState", latent.defLockedState);
  if (latent.defUIPriority !== undefined) {
    attrs.push(`w:defUIPriority="${latent.defUIPriority}"`);
  }
  pushBooleanAttr(attrs, "w:defSemiHidden", latent.defSemiHidden);
  pushBooleanAttr(attrs, "w:defUnhideWhenUsed", latent.defUnhideWhenUsed);
  pushBooleanAttr(attrs, "w:defQFormat", latent.defQFormat);
  return `<w:latentStyles${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}/>`;
};

export const serializeStyle = (style: Style): string => {
  const attrs = [`w:type="${style.type}"`, `w:styleId="${escapeXmlAttribute(style.styleId)}"`];
  if (style.default) {
    attrs.push('w:default="1"');
  }
  const parts = [`<w:name w:val="${escapeXmlAttribute(style.name ?? style.styleId)}"/>`];
  if (style.basedOn) {
    parts.push(`<w:basedOn w:val="${escapeXmlAttribute(style.basedOn)}"/>`);
  }
  if (style.next) {
    parts.push(`<w:next w:val="${escapeXmlAttribute(style.next)}"/>`);
  }
  if (style.link) {
    parts.push(`<w:link w:val="${escapeXmlAttribute(style.link)}"/>`);
  }
  if (style.uiPriority !== undefined) {
    parts.push(`<w:uiPriority w:val="${style.uiPriority}"/>`);
  }
  pushOnOffElement(parts, style.hidden, "hidden");
  pushOnOffElement(parts, style.semiHidden, "semiHidden");
  pushOnOffElement(parts, style.unhideWhenUsed, "unhideWhenUsed");
  pushOnOffElement(parts, style.qFormat, "qFormat");
  pushOnOffElement(parts, style.personal, "personal");
  parts.push(serializeParagraphFormatting(style.pPr));
  parts.push(serializeTextFormatting(style.rPr));
  parts.push(serializeTableFormatting(style.tblPr));
  parts.push(serializeTableRowFormatting(style.trPr));
  parts.push(serializeTableCellFormatting(style.tcPr));
  for (const conditional of style.tblStylePr ?? []) {
    parts.push(
      `<w:tblStylePr w:type="${conditional.type}">${serializeParagraphFormatting(conditional.pPr)}${serializeTextFormatting(conditional.rPr)}${serializeTableFormatting(conditional.tblPr)}${serializeTableRowFormatting(conditional.trPr)}${serializeTableCellFormatting(conditional.tcPr)}</w:tblStylePr>`,
    );
  }
  return `<w:style ${attrs.join(" ")}>${parts.join("")}</w:style>`;
};

const pushBooleanAttr = (attrs: string[], name: string, value: boolean | undefined): void => {
  if (value !== undefined) {
    attrs.push(`${name}="${value ? 1 : 0}"`);
  }
};
