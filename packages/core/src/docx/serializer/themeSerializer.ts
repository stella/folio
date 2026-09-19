import type { Theme, ThemeColorScheme, ThemeFont } from "../../types/document";
import { serializePartElement } from "./partNamespaces";
import { escapeXmlAttribute } from "@stll/docx-core";

export const serializeThemeXml = (theme: Theme): string => {
  const name = escapeXmlAttribute(theme.name ?? "Folio Theme");
  const elements = `${serializeColorScheme(theme.colorScheme)}${serializeFontScheme(theme)}${serializeFormatScheme(theme)}`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    serializePartElement({
      partPath: "word/theme/theme1.xml",
      rootName: "a:theme",
      rootAttributes: `name="${name}"`,
      baselinePrefixes: ["a"],
      sourceBindings: undefined,
      body: `<a:themeElements>${elements}</a:themeElements>`,
    })
  );
};

const serializeColorScheme = (colors: ThemeColorScheme | undefined): string => {
  const values = {
    dk1: colors?.dk1 ?? "000000",
    lt1: colors?.lt1 ?? "FFFFFF",
    dk2: colors?.dk2 ?? "44546A",
    lt2: colors?.lt2 ?? "E7E6E6",
    accent1: colors?.accent1 ?? "4472C4",
    accent2: colors?.accent2 ?? "ED7D31",
    accent3: colors?.accent3 ?? "A5A5A5",
    accent4: colors?.accent4 ?? "FFC000",
    accent5: colors?.accent5 ?? "5B9BD5",
    accent6: colors?.accent6 ?? "70AD47",
    hlink: colors?.hlink ?? "0563C1",
    folHlink: colors?.folHlink ?? "954F72",
  };
  const entries = Object.entries(values)
    .map(
      ([slot, value]) => `<a:${slot}><a:srgbClr val="${escapeXmlAttribute(value)}"/></a:${slot}>`,
    )
    .join("");
  return `<a:clrScheme name="Folio">${entries}</a:clrScheme>`;
};

const serializeFontScheme = (theme: Theme): string => {
  const major = serializeThemeFont(theme.fontScheme?.majorFont, "majorFont", "Arial");
  const minor = serializeThemeFont(theme.fontScheme?.minorFont, "minorFont", "Arial");
  return `<a:fontScheme name="Folio">${major}${minor}</a:fontScheme>`;
};

const serializeThemeFont = (
  font: ThemeFont | undefined,
  element: "majorFont" | "minorFont",
  fallback: string,
): string => {
  const scriptFonts = Object.entries(font?.fonts ?? {})
    .map(
      ([script, typeface]) =>
        `<a:font script="${escapeXmlAttribute(script)}" typeface="${escapeXmlAttribute(typeface)}"/>`,
    )
    .join("");
  return `<a:${element}><a:latin typeface="${escapeXmlAttribute(font?.latin ?? fallback)}"/><a:ea typeface="${escapeXmlAttribute(font?.ea ?? "")}"/><a:cs typeface="${escapeXmlAttribute(font?.cs ?? "")}"/>${scriptFonts}</a:${element}>`;
};

const serializeFormatScheme = (theme: Theme): string => {
  const name = escapeXmlAttribute(theme.formatScheme?.name ?? "Folio");
  return `<a:fmtScheme name="${name}"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>`;
};
