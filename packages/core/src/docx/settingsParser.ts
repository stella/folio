/**
 * settings.xml parser
 *
 * Extracts document-wide settings the layout pipeline consumes at render
 * time. We deliberately read only the handful of settings that affect
 * layout, including the line-breaking compatibility controls Folio supports;
 * the rest of settings.xml is preserved opaquely by the rezip step.
 *
 * See ECMA-376 §17.15 for the full settings part.
 */

import type { DocumentSettings } from "../types/document";
import {
  findChild,
  findChildren,
  getAttribute,
  parseBooleanElement,
  parseXmlDocument,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

export type FolioDocumentSettings = DocumentSettings & {
  /** Swap left/right section margins on even physical pages. */
  mirrorMargins?: boolean;
};

/** OOXML default per §17.6.13 when `w:defaultTabStop` is absent. */
export const DEFAULT_TAB_STOP_TWIPS = 720;

/**
 * Sanity cap on `w:defaultTabStop`. Word's maximum margin is ~22 inches
 * (31 680 twips); anything past that is corruption or a hostile input and
 * we substitute the OOXML default instead.
 */
const MAX_TAB_STOP_TWIPS = 31_680;
const MAX_HYPHENATION_ZONE_TWIPS = 31_680;
const MAX_CONSECUTIVE_HYPHEN_LIMIT = 255;

export function parseSettings(xml: string | null): FolioDocumentSettings {
  const root = xml ? (parseXmlDocument(xml) as XmlElement | null) : null;
  const settings: FolioDocumentSettings = {
    defaultTabStop: parseDefaultTabStop(root),
  };
  // `w:evenAndOddHeaders` lives in settings.xml, not sectPr. Only record the
  // "on" state; absence means odd/even share one header.
  const evenAndOddHeaders = root ? findChild(root, "w", "evenAndOddHeaders") : null;
  if (evenAndOddHeaders && parseBooleanElement(evenAndOddHeaders)) {
    settings.evenAndOddHeaders = true;
  }
  const mirrorMargins = root ? findChild(root, "w", "mirrorMargins") : null;
  if (mirrorMargins && parseBooleanElement(mirrorMargins)) {
    settings.mirrorMargins = true;
  }

  // `w:themeFontLang` selects the concrete typeface for the empty `<a:ea>` /
  // `<a:cs>` theme slots (see eigenpal/docx-editor#949). Record only the
  // EastAsian/bidi tags; the primary `w:val` lang does not drive theme fonts.
  const themeFontLangEl = root ? findChild(root, "w", "themeFontLang") : null;
  const eastAsiaLang = themeFontLangEl
    ? getAttribute(themeFontLangEl, "w", "eastAsia") || undefined
    : undefined;
  const bidiLang = themeFontLangEl
    ? getAttribute(themeFontLangEl, "w", "bidi") || undefined
    : undefined;
  if (eastAsiaLang || bidiLang) {
    settings.themeFontLang = {
      ...(eastAsiaLang ? { eastAsia: eastAsiaLang } : {}),
      ...(bidiLang ? { bidi: bidiLang } : {}),
    };
  }

  const autoHyphenation = parseOnOffSetting(root, "autoHyphenation");
  if (autoHyphenation !== undefined) {
    settings.autoHyphenation = autoHyphenation;
  }
  const doNotHyphenateCaps = parseOnOffSetting(root, "doNotHyphenateCaps");
  if (doNotHyphenateCaps !== undefined) {
    settings.doNotHyphenateCaps = doNotHyphenateCaps;
  }
  const consecutiveHyphenLimit = parseIntegerSetting(
    root,
    "consecutiveHyphenLimit",
    MAX_CONSECUTIVE_HYPHEN_LIMIT,
  );
  if (consecutiveHyphenLimit !== undefined) {
    settings.consecutiveHyphenLimit = consecutiveHyphenLimit;
  }
  const hyphenationZoneTwips = parseIntegerSetting(
    root,
    "hyphenationZone",
    MAX_HYPHENATION_ZONE_TWIPS,
  );
  if (hyphenationZoneTwips !== undefined) {
    settings.hyphenationZoneTwips = hyphenationZoneTwips;
  }

  const noLineBreaksBefore = parseKinsokuOverride(root, "noLineBreaksBefore");
  const noLineBreaksAfter = parseKinsokuOverride(root, "noLineBreaksAfter");
  const compat = root ? findChild(root, "w", "compat") : null;
  const compatibilityMode = parseCompatibilityMode(compat);
  if (compatibilityMode !== undefined) {
    settings.compatibilityMode = compatibilityMode;
  }
  const splitPageBreakAndParagraphMark = compat
    ? findChild(compat, "w", "splitPgBreakAndParaMark")
    : null;
  if (splitPageBreakAndParagraphMark && parseBooleanElement(splitPageBreakAndParagraphMark)) {
    settings.splitPageBreakAndParagraphMark = true;
  }
  const applyBreakingRules = compat ? findChild(compat, "w", "applyBreakingRules") : null;
  const useLegacyEthiopicAmharicRules =
    applyBreakingRules !== null && parseBooleanElement(applyBreakingRules);
  if (noLineBreaksBefore || noLineBreaksAfter || useLegacyEthiopicAmharicRules) {
    settings.lineBreakRules = {
      ...(noLineBreaksBefore ? { noLineBreaksBefore } : {}),
      ...(noLineBreaksAfter ? { noLineBreaksAfter } : {}),
      ...(useLegacyEthiopicAmharicRules ? { useLegacyEthiopicAmharicRules: true } : {}),
    };
  }

  return settings;
}

function parseCompatibilityMode(compat: XmlElement | null): number | undefined {
  if (!compat) {
    return undefined;
  }
  const setting = findChildren(compat, "w", "compatSetting").find(
    (candidate) => getAttribute(candidate, "w", "name") === "compatibilityMode",
  );
  const raw = setting ? getAttribute(setting, "w", "val") : null;
  if (raw === null || !/^\d+$/u.test(raw)) {
    return undefined;
  }
  const parsed = Number.parseInt(raw, 10);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

type OnOffSettingName = "autoHyphenation" | "doNotHyphenateCaps";

function parseOnOffSetting(root: XmlElement | null, name: OnOffSettingName): boolean | undefined {
  if (!root) {
    return undefined;
  }
  const element = findChild(root, "w", name);
  return element ? parseBooleanElement(element) : undefined;
}

type IntegerSettingName = "consecutiveHyphenLimit" | "hyphenationZone";

function parseIntegerSetting(
  root: XmlElement | null,
  name: IntegerSettingName,
  maximum: number,
): number | undefined {
  if (!root) {
    return undefined;
  }
  const element = findChild(root, "w", name);
  const raw = element ? getAttribute(element, "w", "val") : null;
  if (raw === null || !/^\d+$/u.test(raw)) {
    return undefined;
  }
  const parsed = Number.parseInt(raw, 10);
  return Number.isSafeInteger(parsed) && parsed <= maximum ? parsed : undefined;
}

function parseKinsokuOverride(
  root: XmlElement | null,
  name: "noLineBreaksBefore" | "noLineBreaksAfter",
): { language?: string; characters: string } | undefined {
  if (!root) {
    return undefined;
  }
  const element = findChild(root, "w", name);
  const characters = element ? getAttribute(element, "w", "val") : null;
  if (!characters) {
    return undefined;
  }
  const language = getAttribute(element, "w", "lang") || undefined;
  return {
    characters,
    ...(language ? { language } : {}),
  };
}

function parseDefaultTabStop(root: XmlElement | null): number {
  if (!root) {
    return DEFAULT_TAB_STOP_TWIPS;
  }
  const el = findChild(root, "w", "defaultTabStop");
  if (!el) {
    return DEFAULT_TAB_STOP_TWIPS;
  }
  const raw = getAttribute(el, "w", "val");
  if (raw === null) {
    return DEFAULT_TAB_STOP_TWIPS;
  }
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed) || parsed <= 0 || parsed > MAX_TAB_STOP_TWIPS) {
    return DEFAULT_TAB_STOP_TWIPS;
  }
  return parsed;
}
