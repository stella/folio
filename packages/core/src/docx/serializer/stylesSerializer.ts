import { escapeXmlAttribute, pushOnOffElement } from "@stll/docx-core";

import type { Style, StyleDefinitions } from "../../types/document";
import { serializeParagraphFormatting } from "./paragraphSerializer";
import { serializePartElement } from "./partNamespaces";
import {
  serializeTableCellFormatting,
  serializeTableFormatting,
  serializeTableRowFormatting,
} from "./tableSerializer";
import { serializeTextFormatting } from "./textFormattingSerializer";

export const serializeStylesXml = (definitions: StyleDefinitions): string => {
  const docDefaults = serializeDocumentDefaults(definitions);
  const latentStyles = serializeLatentStyles(definitions);
  const styles = definitions.styles
    .map((style) => serializeStyle(style))
    .join("");
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

/** The `w:docDefaults` element a style table's defaults serialize to. */
export const serializeDocumentDefaults = (
  definitions: StyleDefinitions,
): string => {
  const rPr = serializeTextFormatting(definitions.docDefaults?.rPr);
  const rPrDefault =
    rPr === "" ? "<w:rPrDefault/>" : `<w:rPrDefault>${rPr}</w:rPrDefault>`;
  const pPr =
    serializeParagraphFormatting(definitions.docDefaults?.pPr) || "<w:pPr/>";
  return `<w:docDefaults>${rPrDefault}<w:pPrDefault>${pPr}</w:pPrDefault></w:docDefaults>`;
};

type LatentStyles = NonNullable<StyleDefinitions["latentStyles"]>;

const latentStylesAttributes = (latent: LatentStyles): string[] => {
  const attrs: string[] = [];
  pushBooleanAttr(attrs, "w:defLockedState", latent.defLockedState);
  if (latent.defUIPriority !== undefined) {
    attrs.push(`w:defUIPriority="${latent.defUIPriority}"`);
  }
  pushBooleanAttr(attrs, "w:defSemiHidden", latent.defSemiHidden);
  pushBooleanAttr(attrs, "w:defUnhideWhenUsed", latent.defUnhideWhenUsed);
  pushBooleanAttr(attrs, "w:defQFormat", latent.defQFormat);
  return attrs;
};

const serializeLatentStyles = (definitions: StyleDefinitions): string => {
  const latent = definitions.latentStyles;
  if (!latent) {
    return "";
  }
  const attrs = latentStylesAttributes(latent);
  return `<w:latentStyles${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}/>`;
};

/**
 * A `w:latentStyles` element carrying the model's defaults around the
 * exceptions the source element listed, which the model does not hold.
 */
export const serializeLatentStylesAround = (
  latent: LatentStyles,
  exceptions: string,
): string => {
  const attrs = latentStylesAttributes(latent);
  if (latent.count !== undefined) attrs.push(`w:count="${latent.count}"`);
  const open = `<w:latentStyles${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}`;
  return exceptions === ""
    ? `${open}/>`
    : `${open}>${exceptions}</w:latentStyles>`;
};

/** ECMA-376 `CT_Style` child order. */
const STYLE_CHILD_ORDER: readonly string[] = [
  "name",
  "aliases",
  "basedOn",
  "next",
  "link",
  "autoRedefine",
  "hidden",
  "uiPriority",
  "semiHidden",
  "unhideWhenUsed",
  "qFormat",
  "locked",
  "personal",
  "personalCompose",
  "personalReply",
  "rsid",
  "pPr",
  "rPr",
  "tblPr",
  "trPr",
  "tcPr",
  "tblStylePr",
];

/** The `w:style` children and attributes the model holds. */
export const MODELED_STYLE_CHILDREN: ReadonlySet<string> = new Set([
  "name",
  "basedOn",
  "next",
  "link",
  "uiPriority",
  "hidden",
  "semiHidden",
  "unhideWhenUsed",
  "qFormat",
  "personal",
  "pPr",
  "rPr",
  "tblPr",
  "trPr",
  "tcPr",
  "tblStylePr",
]);
export const MODELED_STYLE_ATTRIBUTES: ReadonlySet<string> = new Set([
  "type",
  "styleId",
  "default",
]);

// A child of another namespace has no schema slot and goes last.
const styleChildRank = (localName: string): number => {
  const rank = STYLE_CHILD_ORDER.indexOf(localName);
  return rank < 0 ? STYLE_CHILD_ORDER.length : rank;
};

/**
 * Markup the model does not hold, kept from a source definition the style
 * replaces: its unmodeled attributes, spelled for the written element, and its
 * unmodeled children with their `CT_Style` local name (empty for a child of
 * another namespace).
 */
export type StyleSourceMarkup = {
  attributes: readonly string[];
  children: readonly { localName: string; xml: string }[];
};

export const serializeStyle = (
  style: Style,
  source?: StyleSourceMarkup,
): string => {
  const attrs = [
    `w:type="${style.type}"`,
    `w:styleId="${escapeXmlAttribute(style.styleId)}"`,
  ];
  if (style.default) {
    attrs.push('w:default="1"');
  }
  const named: { localName: string; xml: string }[] = [];
  const push = (localName: string, xml: string) => {
    if (xml !== "") named.push({ localName, xml });
  };
  push(
    "name",
    `<w:name w:val="${escapeXmlAttribute(style.name ?? style.styleId)}"/>`,
  );
  if (style.basedOn) {
    push(
      "basedOn",
      `<w:basedOn w:val="${escapeXmlAttribute(style.basedOn)}"/>`,
    );
  }
  if (style.next) {
    push("next", `<w:next w:val="${escapeXmlAttribute(style.next)}"/>`);
  }
  if (style.link) {
    push("link", `<w:link w:val="${escapeXmlAttribute(style.link)}"/>`);
  }
  if (style.uiPriority !== undefined) {
    push("uiPriority", `<w:uiPriority w:val="${style.uiPriority}"/>`);
  }
  for (const name of [
    "hidden",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "personal",
  ] as const) {
    const parts: string[] = [];
    pushOnOffElement(parts, style[name], name);
    push(name, parts.join(""));
  }
  push("pPr", serializeParagraphFormatting(style.pPr));
  push("rPr", serializeTextFormatting(style.rPr));
  push("tblPr", serializeTableFormatting(style.tblPr));
  push("trPr", serializeTableRowFormatting(style.trPr));
  push("tcPr", serializeTableCellFormatting(style.tcPr));
  for (const conditional of style.tblStylePr ?? []) {
    push(
      "tblStylePr",
      `<w:tblStylePr w:type="${conditional.type}">${serializeParagraphFormatting(conditional.pPr)}${serializeTextFormatting(conditional.rPr)}${serializeTableFormatting(conditional.tblPr)}${serializeTableRowFormatting(conditional.trPr)}${serializeTableCellFormatting(conditional.tcPr)}</w:tblStylePr>`,
    );
  }
  // Each kept source child goes before the first written child that follows
  // it in schema order; the model's own children keep their order.
  for (const child of source?.children ?? []) {
    const rank = styleChildRank(child.localName);
    const at = named.findIndex(
      (written) => styleChildRank(written.localName) > rank,
    );
    named.splice(at < 0 ? named.length : at, 0, child);
  }
  attrs.push(...(source?.attributes ?? []));
  const children = named.map(({ xml }) => xml).join("");
  return `<w:style ${attrs.join(" ")}>${children}</w:style>`;
};

const pushBooleanAttr = (
  attrs: string[],
  name: string,
  value: boolean | undefined,
): void => {
  if (value !== undefined) {
    attrs.push(`${name}="${value ? 1 : 0}"`);
  }
};
