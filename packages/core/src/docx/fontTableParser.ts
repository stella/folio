import type { FontInfo, FontTable } from "../types/document";
import { findChild, findChildren, getAttribute, parseXmlDocument } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const FONT_FAMILIES = ["decorative", "modern", "roman", "script", "swiss", "auto"] as const;
const FONT_PITCHES = ["default", "fixed", "variable"] as const;

export const parseFontTable = (xml: string | null | undefined): FontTable | undefined => {
  if (!xml) {
    return undefined;
  }
  const root = parseXmlDocument(xml);
  if (!root) {
    return undefined;
  }
  const fonts: FontInfo[] = [];
  for (const element of findChildren(root, "w", "font")) {
    const font = parseFont(element);
    if (font) {
      fonts.push(font);
    }
  }
  return fonts.length > 0 ? { fonts } : undefined;
};

const parseFont = (element: XmlElement): FontInfo | undefined => {
  const name = getAttribute(element, "w", "name");
  if (!name) {
    return undefined;
  }
  const font: FontInfo = { name };
  assignValue(font, "altName", childValue(element, "altName"));
  assignValue(font, "panose1", childValue(element, "panose1"));
  assignValue(font, "charset", childValue(element, "charset"));

  const family = childValue(element, "family");
  if (family && isFontFamily(family)) {
    font.family = family;
  }
  const pitch = childValue(element, "pitch");
  if (pitch && isFontPitch(pitch)) {
    font.pitch = pitch;
  }

  const sig = findChild(element, "w", "sig");
  if (sig) {
    const parsed: NonNullable<FontInfo["sig"]> = {};
    assignSignatureValue(parsed, "usb0", getAttribute(sig, "w", "usb0") ?? undefined);
    assignSignatureValue(parsed, "usb1", getAttribute(sig, "w", "usb1") ?? undefined);
    assignSignatureValue(parsed, "usb2", getAttribute(sig, "w", "usb2") ?? undefined);
    assignSignatureValue(parsed, "usb3", getAttribute(sig, "w", "usb3") ?? undefined);
    assignSignatureValue(parsed, "csb0", getAttribute(sig, "w", "csb0") ?? undefined);
    assignSignatureValue(parsed, "csb1", getAttribute(sig, "w", "csb1") ?? undefined);
    if (Object.values(parsed).some((value) => value !== undefined)) {
      font.sig = parsed;
    }
  }

  assignValue(font, "embedRegular", relationshipId(element, "embedRegular"));
  assignValue(font, "embedBold", relationshipId(element, "embedBold"));
  assignValue(font, "embedItalic", relationshipId(element, "embedItalic"));
  assignValue(font, "embedBoldItalic", relationshipId(element, "embedBoldItalic"));
  return font;
};

const childValue = (element: XmlElement, name: string): string | undefined => {
  const child = findChild(element, "w", name);
  return child ? (getAttribute(child, "w", "val") ?? undefined) : undefined;
};

const relationshipId = (element: XmlElement, name: string): string | undefined => {
  const child = findChild(element, "w", name);
  return child ? (getAttribute(child, "r", "id") ?? undefined) : undefined;
};

const assignValue = <Key extends keyof FontInfo>(
  font: FontInfo,
  key: Key,
  value: FontInfo[Key] | undefined,
): void => {
  if (value !== undefined) {
    font[key] = value;
  }
};

const assignSignatureValue = <Key extends keyof NonNullable<FontInfo["sig"]>>(
  signature: NonNullable<FontInfo["sig"]>,
  key: Key,
  value: string | undefined,
): void => {
  if (value !== undefined) {
    signature[key] = value;
  }
};

const isFontFamily = (value: string): value is NonNullable<FontInfo["family"]> =>
  FONT_FAMILIES.some((family) => family === value);

const isFontPitch = (value: string): value is NonNullable<FontInfo["pitch"]> =>
  FONT_PITCHES.some((pitch) => pitch === value);
