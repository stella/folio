import type { Theme, ThemeColorScheme, ThemeFont } from "../../types/document";
import { escapeXml } from "./xmlUtils";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export const serializeThemeXml = (theme: Theme): string => {
  const name = escapeXml(theme.name ?? "Folio Theme");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="${A_NS}" name="${name}"><a:themeElements>${serializeColorScheme(theme.colorScheme)}${serializeFontScheme(theme)}${serializeFormatScheme(theme)}</a:themeElements></a:theme>`;
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
    .map(([slot, value]) => `<a:${slot}><a:srgbClr val="${escapeXml(value)}"/></a:${slot}>`)
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
        `<a:font script="${escapeXml(script)}" typeface="${escapeXml(typeface)}"/>`,
    )
    .join("");
  return `<a:${element}><a:latin typeface="${escapeXml(font?.latin ?? fallback)}"/><a:ea typeface="${escapeXml(font?.ea ?? "")}"/><a:cs typeface="${escapeXml(font?.cs ?? "")}"/>${scriptFonts}</a:${element}>`;
};

const serializeFormatScheme = (theme: Theme): string => {
  const name = escapeXml(theme.formatScheme?.name ?? "Folio");
  return `<a:fmtScheme name="${name}"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>`;
};
