/**
 * `folio lang`: the document's proofing language. Reports the `w:lang`
 * values every part states and, given a language, writes it where a
 * consumer looks first: the run-property defaults in `w:docDefaults`, the
 * default paragraph style, every `w:lang` a style already states, and
 * `w:themeFontLang` in settings. With `runs`, every `w:lang` a run or a
 * paragraph mark states in the body, headers, footers, and notes is set
 * too; runs that state none inherit the defaults.
 */

import { Result } from "better-result";

import { SEQUENCE_CHILDREN, SETTINGS_CHILDREN } from "@stll/docx-core/schema";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import {
  childElements,
  ensureWChild,
  isW,
  mainDocumentPart,
  type OoxmlPackage,
  RELATIONSHIP_TYPES,
  setWAttribute,
  storyParts,
  walkElements,
  wAttribute,
  type XmlElement,
} from "./ooxml-package";

/** The `w:lang` attributes, each with the argument that sets it. */
const LANGUAGE_ARGUMENTS = [
  ["val", "set"],
  ["eastAsia", "eastAsia"],
  ["bidi", "bidi"],
] as const;

type LanguageAttribute = (typeof LANGUAGE_ARGUMENTS)[number][0];

export type LanguageSetting = Partial<Record<LanguageAttribute, string>>;

/** Language tags each `w:lang` attribute holds in one part, with how often. */
export type PartLanguages = {
  part: string;
} & Partial<Record<LanguageAttribute, Record<string, number>>>;

export type ProofingLanguageReport = {
  /** The languages each part states, before any change. */
  found: PartLanguages[];
  /** What was set, when a language was given. */
  set?: LanguageSetting & {
    docDefaults: boolean;
    styles: number;
    themeFontLang: boolean;
    runs: number;
  };
};

/** A BCP 47 tag, or the hexadecimal language id the schema also admits. */
const LANGUAGE_TAG =
  /^(?:[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*|[0-9A-Fa-f]{4})$/u;

const STYLES_CHILDREN = ["docDefaults", "latentStyles", "style"] as const;
const DOC_DEFAULTS_CHILDREN = ["rPrDefault", "pPrDefault"] as const;
const STYLE_CHILDREN = [
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
] as const;
const RUN_PROPERTIES = SEQUENCE_CHILDREN["run-properties"];

const invalidInput = (message: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message });

const settingFrom = (
  args: Readonly<Record<string, unknown>>,
): Result<LanguageSetting | null, FolioCliError> => {
  const setting: LanguageSetting = {};
  for (const [attribute, key] of LANGUAGE_ARGUMENTS) {
    const value = args[key];
    if (value === undefined) continue;
    if (typeof value !== "string" || !LANGUAGE_TAG.test(value)) {
      return Result.err(
        invalidInput(
          `${key} must be a language tag such as cs-CZ or en-GB; got ${JSON.stringify(value)}.`,
        ),
      );
    }
    setting[attribute] = value;
  }
  if (Object.keys(setting).length > 0) return Result.ok(setting);
  return args["runs"] === true
    ? Result.err(invalidInput("runs needs a language to set."))
    : Result.ok(null);
};

/** Count the languages every `w:lang` (and `w:themeFontLang`) below `root` states. */
const languagesOf = (part: string, root: XmlElement): PartLanguages | null => {
  const counts: Partial<Record<LanguageAttribute, Record<string, number>>> = {};
  let found = false;
  walkElements(root, (element) => {
    if (!isW(element, "lang") && !isW(element, "themeFontLang")) return;
    for (const [attribute] of LANGUAGE_ARGUMENTS) {
      const value = wAttribute(element, attribute);
      if (value === undefined) continue;
      const tally = (counts[attribute] ??= {});
      tally[value] = (tally[value] ?? 0) + 1;
      found = true;
    }
  });
  return found ? { part, ...counts } : null;
};

/** Write the setting's attributes onto a `w:lang`; whether any value changed. */
const applySetting = (
  element: XmlElement,
  setting: LanguageSetting,
): boolean => {
  let changed = false;
  for (const [attribute] of LANGUAGE_ARGUMENTS) {
    const value = setting[attribute];
    if (value === undefined || wAttribute(element, attribute) === value)
      continue;
    setWAttribute(element, attribute, value);
    changed = true;
  }
  return changed;
};

/** The `w:lang` of a run-properties element, created in order when missing. */
const langOf = (runProperties: XmlElement): XmlElement =>
  ensureWChild(runProperties, "lang", RUN_PROPERTIES);

/** The part a main-document relationship of `type` names. */
const relatedPart = (
  pkg: OoxmlPackage,
  type: string,
): Result<string | null, FolioCliError> => {
  const main = mainDocumentPart(pkg);
  if (main.isErr()) return Result.err(main.error);
  if (main.value === null) return Result.ok(null);
  return pkg
    .relationships(main.value)
    .map(
      (relationships) =>
        relationships.find((relationship) => relationship.type === type)
          ?.part ?? null,
    );
};

/** Set the language in the styles part: whether the defaults changed, and how many styles did. */
const setStyles = (
  styles: XmlElement,
  setting: LanguageSetting,
): { docDefaults: boolean; styles: number } => {
  const docDefaults = ensureWChild(styles, "docDefaults", STYLES_CHILDREN);
  const rPrDefault = ensureWChild(
    docDefaults,
    "rPrDefault",
    DOC_DEFAULTS_CHILDREN,
  );
  const defaults = applySetting(
    langOf(ensureWChild(rPrDefault, "rPr", ["rPr"])),
    setting,
  );
  let changed = 0;
  for (const style of childElements(styles)) {
    if (!isW(style, "style")) continue;
    const isDefaultParagraph =
      wAttribute(style, "type") === "paragraph" &&
      ["1", "true", "on"].includes(wAttribute(style, "default") ?? "");
    const runProperties = childElements(style).find((child) =>
      isW(child, "rPr"),
    );
    const stated = runProperties
      ? childElements(runProperties).find((child) => isW(child, "lang"))
      : undefined;
    if (stated === undefined && !isDefaultParagraph) continue;
    const lang = stated ?? langOf(ensureWChild(style, "rPr", STYLE_CHILDREN));
    if (applySetting(lang, setting)) changed += 1;
  }
  return { docDefaults: defaults, styles: changed };
};

/** Set every `w:lang` a run or paragraph mark states in a story; how many. */
const setRuns = (story: XmlElement, setting: LanguageSetting): number => {
  let changed = 0;
  walkElements(story, (element) => {
    if (!isW(element, "rPr")) return;
    for (const lang of childElements(element)) {
      if (isW(lang, "lang") && applySetting(lang, setting)) changed += 1;
    }
  });
  return changed;
};

/** Report the languages in every part and, given one, set it. */
export const setProofingLanguage = (
  pkg: OoxmlPackage,
  args: Readonly<Record<string, unknown>>,
): Result<ProofingLanguageReport, FolioCliError> => {
  const setting = settingFrom(args);
  if (setting.isErr()) return Result.err(setting.error);

  const stylesPart = relatedPart(pkg, RELATIONSHIP_TYPES.styles);
  if (stylesPart.isErr()) return Result.err(stylesPart.error);
  const settingsPart = relatedPart(pkg, RELATIONSHIP_TYPES.settings);
  if (settingsPart.isErr()) return Result.err(settingsPart.error);
  const stories = storyParts(pkg);
  if (stories.isErr()) return Result.err(stories.error);

  const parts = [
    ...(stylesPart.value === null ? [] : [stylesPart.value]),
    ...(settingsPart.value === null ? [] : [settingsPart.value]),
    ...stories.value,
  ];
  const found: PartLanguages[] = [];
  const roots = new Map<string, XmlElement>();
  for (const part of parts) {
    const root = pkg.root(part);
    if (root.isErr()) return Result.err(root.error);
    if (root.value === null) continue;
    roots.set(part, root.value);
    const languages = languagesOf(part, root.value);
    if (languages !== null) found.push(languages);
  }
  if (setting.value === null) return Result.ok({ found });

  const applied = {
    ...setting.value,
    docDefaults: false,
    styles: 0,
    themeFontLang: false,
    runs: 0,
  };
  const styles =
    stylesPart.value === null ? undefined : roots.get(stylesPart.value);
  if (stylesPart.value !== null && styles !== undefined) {
    const changed = setStyles(styles, setting.value);
    Object.assign(applied, changed);
    if (changed.docDefaults || changed.styles > 0) pkg.touch(stylesPart.value);
  }
  const settings =
    settingsPart.value === null ? undefined : roots.get(settingsPart.value);
  if (settingsPart.value !== null && settings !== undefined) {
    applied.themeFontLang = applySetting(
      ensureWChild(settings, "themeFontLang", SETTINGS_CHILDREN),
      setting.value,
    );
    if (applied.themeFontLang) pkg.touch(settingsPart.value);
  }
  if (args["runs"] === true) {
    for (const part of stories.value) {
      const story = roots.get(part);
      if (story === undefined) continue;
      const changed = setRuns(story, setting.value);
      if (changed > 0) pkg.touch(part);
      applied.runs += changed;
    }
  }
  return Result.ok({ found, set: applied });
};
