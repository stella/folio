import { panic } from "better-result";
import type {
  AuthoredParagraphProperties,
  AuthoredParagraphPropertyKey,
  NonEmptyNumberingLevelIndentGeometry,
  NumberingLevelIndentGeometry,
  NumberingLevelIndentProvenance,
  ParagraphMarkProperties,
} from "@stll/docx-core/model";
import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  selectAuthoredParagraphProperties,
} from "@stll/docx-core/model";

import {
  ParagraphPropertySourceValidationError,
  ParagraphPropertyTransientTemplateHandle,
  type ParagraphPropertySourceAttribute,
} from "../docx/paragraphPropertySourceIdentity";
import {
  FRAME_WRAP_VALUES,
  FRAME_X_ALIGN_VALUES,
  FRAME_Y_ALIGN_VALUES,
  EMPHASIS_MARK_VALUES,
  FONT_HINT_VALUES,
  FONT_THEME_VALUES,
  HIGHLIGHT_COLOR_VALUES,
  LINE_SPACING_RULE_VALUES,
  PARAGRAPH_ALIGNMENT_VALUES,
  SHADING_PATTERN_VALUES,
  TAB_LEADER_VALUES,
  TAB_STOP_ALIGNMENT_VALUES,
  TEXT_EFFECT_VALUES,
  THEME_COLOR_SLOT_VALUES,
  UNDERLINE_STYLE_VALUES,
} from "../types/documentEnumValues";
import type {
  BorderSpec,
  ColorValue,
  ParagraphFormatting,
  ShadingProperties,
  TabStop,
  TextFormatting,
} from "../types/document";
import {
  PARAGRAPH_SPACING_INHERITANCE_SOURCE,
  type ParagraphMarkEffectiveProperties,
  type ParagraphMarkProjectionContext,
  type ParagraphSpacingInheritance,
  type SerializedParagraphPropertyProjectionContext,
} from "./paragraphPropertyContext";

type SerializedParagraphPropertyStateCommon = {
  authoredPPr: AuthoredParagraphProperties;
  context: SerializedParagraphPropertyProjectionContext;
};

export type SerializedPersistableParagraphPropertyState =
  | (SerializedParagraphPropertyStateCommon & {
      type: "imported";
      token: string;
    })
  | (SerializedParagraphPropertyStateCommon & {
      type: "editor-created";
    });

/** PM/Yjs wire data is persistable by construction; transient handles have no wire branch. */
export type SerializedParagraphPropertyState = SerializedPersistableParagraphPropertyState;

type DeepReadonly<Value> = Value extends (...args: never[]) => unknown
  ? Value
  : Value extends readonly (infer Entry)[]
    ? readonly DeepReadonly<Entry>[]
    : Value extends object
      ? { readonly [Key in keyof Value]: DeepReadonly<Value[Key]> }
      : Value;

export type ParagraphPropertyProjectionContext =
  DeepReadonly<SerializedParagraphPropertyProjectionContext>;

const TRUSTED_PARAGRAPH_PROPERTY_STATE: unique symbol = Symbol("trusted-paragraph-property-state");

type TrustedParagraphPropertyState = {
  readonly [TRUSTED_PARAGRAPH_PROPERTY_STATE]: true;
  readonly authoredPPr: DeepReadonly<AuthoredParagraphProperties>;
  readonly context: ParagraphPropertyProjectionContext;
};

export type PersistableParagraphPropertyState = (
  | {
      readonly type: "imported";
      readonly token: string;
    }
  | {
      readonly type: "editor-created";
    }
) &
  TrustedParagraphPropertyState;

/** Trusted state used after the serialized PM/Yjs boundary has been validated. */
export type ParagraphPropertyState =
  | PersistableParagraphPropertyState
  | ({
      readonly type: "transient-template";
      readonly handle: ParagraphPropertyTransientTemplateHandle;
    } & TrustedParagraphPropertyState);

export type ParagraphPropertyMutation = {
  [Key in AuthoredParagraphPropertyKey]:
    | { key: Key; mutation: { type: "remove" } }
    | {
        key: Key;
        mutation: {
          type: "set";
          value: Exclude<AuthoredParagraphProperties[Key], undefined>;
        };
      };
}[AuthoredParagraphPropertyKey];

export type ParagraphMarkPropertyMutation =
  | {
      key: "runProperties";
      mutation: { type: "remove" } | { type: "set"; value: TextFormatting };
    }
  | {
      key: "runInWithNext";
      mutation: { type: "remove" } | { type: "set"; value: boolean };
    };

export type ParagraphPropertyAuthoredTransition =
  | { type: "preserve" }
  | { type: "mutate"; mutations: readonly ParagraphPropertyMutation[] }
  | { type: "replace"; authoredPPr: AuthoredParagraphProperties };

export type ParagraphPropertyContextTransition =
  | { type: "preserve" }
  | { type: "replace"; context: SerializedParagraphPropertyProjectionContext }
  | {
      type: "mutate-paragraph-mark";
      authored: readonly ParagraphMarkPropertyMutation[];
      effective: readonly ParagraphMarkEffectivePropertyMutation[];
    };

export type ParagraphMarkEffectivePropertyMutation =
  | {
      key: "defaultTextFormatting";
      mutation: { type: "remove" } | { type: "set"; value: TextFormatting };
    }
  | {
      key: "runInWithNext";
      mutation: { type: "remove" } | { type: "set"; value: boolean };
    };

export type ParagraphPropertyStateTransition = {
  type: "update";
  authored: ParagraphPropertyAuthoredTransition;
  context: ParagraphPropertyContextTransition;
};

export type ParagraphPropertySplitTransition = {
  type: "split-left-created-right-retains";
};

export type ParagraphPropertyJoinTransition =
  | { type: "join-left-paragraph-mark-retains" }
  | { type: "join-right-paragraph-mark-retains" };

type ValidationResult = { valid: true } | { valid: false };

const INVALID: ValidationResult = { valid: false };
const VALID: ValidationResult = { valid: true };

const PARAGRAPH_PROPERTY_STATE_MAX_DEPTH = 16;
const PARAGRAPH_PROPERTY_STATE_MAX_NODES = 4_096;
const PARAGRAPH_PROPERTY_STATE_MAX_LIST_ENTRIES = 1_024;
const PARAGRAPH_PROPERTY_STATE_MAX_STRING_CODE_UNITS = 16_384;

const assertParagraphPropertyStateBudget = (raw: unknown): void => {
  const pending: { depth: number; value: unknown }[] = [{ depth: 0, value: raw }];
  const seen = new WeakSet<object>();
  let nodeCount = 0;
  while (pending.length > 0) {
    const entry = pending.pop();
    if (!entry) {
      break;
    }
    nodeCount += 1;
    if (
      nodeCount > PARAGRAPH_PROPERTY_STATE_MAX_NODES ||
      entry.depth > PARAGRAPH_PROPERTY_STATE_MAX_DEPTH
    ) {
      throw new ParagraphPropertySourceValidationError({
        code: "state_capacity_exceeded",
        message: "Paragraph-property state complexity capacity was exceeded.",
      });
    }
    if (typeof entry.value === "string") {
      if (entry.value.length > PARAGRAPH_PROPERTY_STATE_MAX_STRING_CODE_UNITS) {
        throw new ParagraphPropertySourceValidationError({
          code: "state_capacity_exceeded",
          message: "Paragraph-property state string capacity was exceeded.",
        });
      }
      continue;
    }
    if (typeof entry.value !== "object" || entry.value === null) {
      continue;
    }
    if (seen.has(entry.value)) {
      continue;
    }
    seen.add(entry.value);
    const keys = Reflect.ownKeys(entry.value);
    if (
      keys.length > PARAGRAPH_PROPERTY_STATE_MAX_LIST_ENTRIES ||
      keys.some((key) => typeof key !== "string")
    ) {
      throw new ParagraphPropertySourceValidationError({
        code: "state_capacity_exceeded",
        message: "Paragraph-property state collection capacity was exceeded.",
      });
    }
    const descriptors = Object.getOwnPropertyDescriptors(entry.value);
    for (const key of keys) {
      if (typeof key !== "string") {
        continue;
      }
      const descriptor = descriptors[key];
      if (!descriptor || !("value" in descriptor)) {
        throw new ParagraphPropertySourceValidationError({
          code: "invalid_state",
          message: "Paragraph-property state cannot contain accessors.",
        });
      }
      pending.push({ depth: entry.depth + 1, value: descriptor.value });
    }
  }
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const hasOnlyKeys = (value: Record<string, unknown>, allowed: ReadonlySet<string>): boolean =>
  Object.keys(value).every((key) => allowed.has(key));

const isFiniteNumber = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value);

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  Number.isSafeInteger(value) && typeof value === "number" && value >= 0;

const isOneOf = <Value extends string>(value: unknown, allowed: readonly Value[]): value is Value =>
  typeof value === "string" && allowed.includes(value as Value);

const validateOptionalBooleanRecord = (
  value: unknown,
  keys: readonly string[],
): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, new Set(keys))) {
    return INVALID;
  }
  return Object.values(value).every((entry) => typeof entry === "boolean") ? VALID : INVALID;
};

const COLOR_VALUE_VALIDATION = {
  rgb: "string",
  auto: "boolean",
  themeColor: "theme-color",
  themeTint: "string",
  themeShade: "string",
} as const satisfies Record<keyof ColorValue, "boolean" | "string" | "theme-color">;
const COLOR_VALUE_KEYS = new Set<string>(Object.keys(COLOR_VALUE_VALIDATION));

const validateColorValue = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, COLOR_VALUE_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(COLOR_VALUE_VALIDATION)) {
    const property = value[key];
    if (property === undefined) {
      continue;
    }
    if (validation === "boolean" && typeof property !== "boolean") {
      return INVALID;
    }
    if (validation === "string" && typeof property !== "string") {
      return INVALID;
    }
    if (validation === "theme-color" && !isOneOf(property, THEME_COLOR_SLOT_VALUES)) {
      return INVALID;
    }
  }
  return VALID;
};

const BORDER_VALIDATION = {
  style: "required-string",
  size: "number",
  space: "number",
  color: "color",
  shadow: "boolean",
  frame: "boolean",
  artRelationshipId: "string",
  topLeftArtRelationshipId: "string",
  topRightArtRelationshipId: "string",
  bottomLeftArtRelationshipId: "string",
  bottomRightArtRelationshipId: "string",
} as const satisfies Record<
  keyof BorderSpec,
  "boolean" | "color" | "number" | "required-string" | "string"
>;
const BORDER_KEYS = new Set<string>(Object.keys(BORDER_VALIDATION));

const validateBorder = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, BORDER_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(BORDER_VALIDATION)) {
    const property = value[key];
    if (validation === "required-string") {
      if (typeof property !== "string") {
        return INVALID;
      }
      continue;
    }
    if (property === undefined) {
      continue;
    }
    if (
      (validation === "string" && typeof property !== "string") ||
      (validation === "number" && !isFiniteNumber(property)) ||
      (validation === "boolean" && typeof property !== "boolean") ||
      (validation === "color" && !validateColorValue(property).valid)
    ) {
      return INVALID;
    }
  }
  return VALID;
};

type ParagraphBorders = NonNullable<AuthoredParagraphProperties["borders"]>;
const PARAGRAPH_BORDER_VALIDATION = {
  top: "border",
  bottom: "border",
  left: "border",
  right: "border",
  between: "border",
  bar: "border",
} as const satisfies Record<keyof ParagraphBorders, "border">;
const PARAGRAPH_BORDER_KEYS = new Set<string>(Object.keys(PARAGRAPH_BORDER_VALIDATION));

const validateBorders = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, PARAGRAPH_BORDER_KEYS)) {
    return INVALID;
  }
  return Object.values(value).every((entry) => validateBorder(entry).valid) ? VALID : INVALID;
};

const SHADING_VALIDATION = {
  color: "color",
  fill: "color",
  pattern: "pattern",
} as const satisfies Record<keyof ShadingProperties, "color" | "pattern">;
const SHADING_KEYS = new Set<string>(Object.keys(SHADING_VALIDATION));

const validateShading = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, SHADING_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(SHADING_VALIDATION)) {
    const property = value[key];
    if (property === undefined) {
      continue;
    }
    if (
      (validation === "color" && !validateColorValue(property).valid) ||
      (validation === "pattern" && !isOneOf(property, SHADING_PATTERN_VALUES))
    ) {
      return INVALID;
    }
  }
  return VALID;
};

const TEXT_FORMATTING_VALIDATION = {
  bold: "boolean",
  boldCs: "boolean",
  italic: "boolean",
  italicCs: "boolean",
  underline: "underline",
  strike: "boolean",
  doubleStrike: "boolean",
  vertAlign: "vertical-alignment",
  smallCaps: "boolean",
  allCaps: "boolean",
  hidden: "boolean",
  color: "color",
  highlight: "highlight",
  shading: "shading",
  fontSize: "number",
  fontSizeCs: "number",
  fontFamily: "font-family",
  language: "language",
  spacing: "number",
  position: "number",
  scale: "number",
  kerning: "number",
  effect: "effect",
  emphasisMark: "emphasis-mark",
  emboss: "boolean",
  imprint: "boolean",
  outline: "boolean",
  shadow: "boolean",
  rtl: "boolean",
  cs: "boolean",
  styleId: "string",
} as const satisfies Record<
  keyof TextFormatting,
  | "boolean"
  | "color"
  | "effect"
  | "emphasis-mark"
  | "font-family"
  | "highlight"
  | "language"
  | "number"
  | "shading"
  | "string"
  | "underline"
  | "vertical-alignment"
>;
const TEXT_FORMATTING_KEYS = new Set<string>(Object.keys(TEXT_FORMATTING_VALIDATION));

type TextUnderline = NonNullable<TextFormatting["underline"]>;
const TEXT_UNDERLINE_VALIDATION = {
  style: "required-style",
  color: "color",
} as const satisfies Record<keyof TextUnderline, "color" | "required-style">;
const TEXT_UNDERLINE_KEYS = new Set<string>(Object.keys(TEXT_UNDERLINE_VALIDATION));

type TextFontFamily = NonNullable<TextFormatting["fontFamily"]>;
const TEXT_FONT_FAMILY_VALIDATION = {
  ascii: "string",
  hAnsi: "string",
  eastAsia: "string",
  cs: "string",
  hint: "font-hint",
  asciiTheme: "font-theme",
  hAnsiTheme: "string",
  eastAsiaTheme: "string",
  csTheme: "string",
} as const satisfies Record<keyof TextFontFamily, "font-hint" | "font-theme" | "string">;
const TEXT_FONT_FAMILY_KEYS = new Set<string>(Object.keys(TEXT_FONT_FAMILY_VALIDATION));

type TextLanguage = NonNullable<TextFormatting["language"]>;
const TEXT_LANGUAGE_VALIDATION = {
  val: "string",
  eastAsia: "string",
  bidi: "string",
} as const satisfies Record<keyof TextLanguage, "string">;
const TEXT_LANGUAGE_KEYS = new Set<string>(Object.keys(TEXT_LANGUAGE_VALIDATION));

const validateTextFormatting = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, TEXT_FORMATTING_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(TEXT_FORMATTING_VALIDATION)) {
    const property = value[key];
    if (property === undefined) {
      continue;
    }
    switch (validation) {
      case "boolean":
        if (typeof property !== "boolean") return INVALID;
        break;
      case "number":
        if (!isFiniteNumber(property)) return INVALID;
        break;
      case "string":
        if (typeof property !== "string") return INVALID;
        break;
      case "vertical-alignment":
        if (!isOneOf(property, ["baseline", "superscript", "subscript"])) return INVALID;
        break;
      case "effect":
        if (!isOneOf(property, TEXT_EFFECT_VALUES)) return INVALID;
        break;
      case "emphasis-mark":
        if (!isOneOf(property, EMPHASIS_MARK_VALUES)) return INVALID;
        break;
      case "highlight":
        if (!isOneOf(property, HIGHLIGHT_COLOR_VALUES)) return INVALID;
        break;
      case "color":
        if (!validateColorValue(property).valid) return INVALID;
        break;
      case "shading":
        if (!validateShading(property).valid) return INVALID;
        break;
      case "underline": {
        if (!isRecord(property) || !hasOnlyKeys(property, TEXT_UNDERLINE_KEYS)) return INVALID;
        for (const [underlineKey, underlineValidation] of Object.entries(
          TEXT_UNDERLINE_VALIDATION,
        )) {
          const underlineProperty = property[underlineKey];
          if (
            (underlineValidation === "required-style" &&
              !isOneOf(underlineProperty, UNDERLINE_STYLE_VALUES)) ||
            (underlineValidation === "color" &&
              underlineProperty !== undefined &&
              !validateColorValue(underlineProperty).valid)
          ) {
            return INVALID;
          }
        }
        break;
      }
      case "font-family": {
        if (!isRecord(property) || !hasOnlyKeys(property, TEXT_FONT_FAMILY_KEYS)) return INVALID;
        for (const [fontKey, fontValidation] of Object.entries(TEXT_FONT_FAMILY_VALIDATION)) {
          const fontProperty = property[fontKey];
          if (fontProperty === undefined) continue;
          if (
            (fontValidation === "string" && typeof fontProperty !== "string") ||
            (fontValidation === "font-hint" && !isOneOf(fontProperty, FONT_HINT_VALUES)) ||
            (fontValidation === "font-theme" && !isOneOf(fontProperty, FONT_THEME_VALUES))
          ) {
            return INVALID;
          }
        }
        break;
      }
      case "language": {
        if (!isRecord(property) || !hasOnlyKeys(property, TEXT_LANGUAGE_KEYS)) return INVALID;
        for (const [languageKey] of Object.entries(TEXT_LANGUAGE_VALIDATION)) {
          const languageProperty = property[languageKey];
          if (languageProperty !== undefined && typeof languageProperty !== "string") {
            return INVALID;
          }
        }
        break;
      }
      default: {
        const exhaustive: never = validation;
        return exhaustive;
      }
    }
  }
  return VALID;
};

export const readTextFormatting = (
  raw: unknown,
): ParagraphPropertySourceAttribute<TextFormatting> => {
  if (raw === null || raw === undefined) {
    return { status: "absent" };
  }
  if (!validateTextFormatting(raw).valid) {
    return { raw, status: "invalid" };
  }
  // SAFETY: the total validator established every TextFormatting field.
  return { status: "valid", value: cloneJsonValue(raw) as TextFormatting };
};

const PARAGRAPH_MARK_PROPERTY_VALIDATION = {
  runProperties: "text-formatting",
  runInWithNext: "boolean",
} as const satisfies Record<keyof ParagraphMarkProperties, "boolean" | "text-formatting">;
const PARAGRAPH_MARK_PROPERTY_KEYS = new Set<string>(
  Object.keys(PARAGRAPH_MARK_PROPERTY_VALIDATION),
);

export const readParagraphMarkProperties = (
  raw: unknown,
): ParagraphPropertySourceAttribute<ParagraphMarkProperties> => {
  if (raw === null || raw === undefined) {
    return { status: "absent" };
  }
  if (!isRecord(raw) || !hasOnlyKeys(raw, PARAGRAPH_MARK_PROPERTY_KEYS)) {
    return { raw, status: "invalid" };
  }
  for (const [key, validation] of Object.entries(PARAGRAPH_MARK_PROPERTY_VALIDATION)) {
    const value = raw[key];
    if (value === undefined) {
      continue;
    }
    if (
      (validation === "boolean" && typeof value !== "boolean") ||
      (validation === "text-formatting" && !validateTextFormatting(value).valid)
    ) {
      return { raw, status: "invalid" };
    }
  }
  // SAFETY: the exact-key checks and nested validator established the mark shape.
  return { status: "valid", value: cloneJsonValue(raw) as ParagraphMarkProperties };
};

const TAB_VALIDATION = {
  position: "required-number",
  alignment: "required-alignment",
  leader: "leader",
} as const satisfies Record<keyof TabStop, "leader" | "required-alignment" | "required-number">;
const TAB_KEYS = new Set<string>(Object.keys(TAB_VALIDATION));

const validateTabs = (value: unknown): ValidationResult => {
  if (!Array.isArray(value)) {
    return INVALID;
  }
  for (const tab of value) {
    if (!isRecord(tab) || !hasOnlyKeys(tab, TAB_KEYS)) {
      return INVALID;
    }
    for (const [key, validation] of Object.entries(TAB_VALIDATION)) {
      const property = tab[key];
      if (validation === "required-number" && !isFiniteNumber(property)) {
        return INVALID;
      }
      if (validation === "required-alignment" && !isOneOf(property, TAB_STOP_ALIGNMENT_VALUES)) {
        return INVALID;
      }
      if (
        validation === "leader" &&
        property !== undefined &&
        !isOneOf(property, TAB_LEADER_VALUES)
      ) {
        return INVALID;
      }
    }
  }
  return VALID;
};

type ParagraphNumberingProperties = NonNullable<AuthoredParagraphProperties["numPr"]>;
const NUM_PR_VALIDATION = {
  numId: "non-negative-integer",
  ilvl: "non-negative-integer",
} as const satisfies Record<keyof ParagraphNumberingProperties, "non-negative-integer">;
const NUM_PR_KEYS = new Set<string>(Object.keys(NUM_PR_VALIDATION));

type ParagraphSpacingExplicit = NonNullable<AuthoredParagraphProperties["spacingExplicit"]>;
const SPACING_EXPLICIT_VALIDATION = {
  before: "boolean",
  after: "boolean",
} as const satisfies Record<keyof ParagraphSpacingExplicit, "boolean">;
const SPACING_EXPLICIT_KEY_LIST = Object.keys(SPACING_EXPLICIT_VALIDATION);

const validateNumPr = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, NUM_PR_KEYS)) {
    return INVALID;
  }
  for (const key of Object.keys(NUM_PR_VALIDATION)) {
    if (value[key] !== undefined && !isNonNegativeSafeInteger(value[key])) {
      return INVALID;
    }
  }
  return VALID;
};

type ParagraphFrameProperties = NonNullable<ParagraphFormatting["frame"]>;
const FRAME_VALIDATION = {
  dropCap: "drop-cap",
  lines: "number",
  width: "number",
  height: "number",
  hSpace: "number",
  vSpace: "number",
  hAnchor: "horizontal-anchor",
  vAnchor: "vertical-anchor",
  x: "number",
  y: "number",
  xAlign: "horizontal-alignment",
  yAlign: "vertical-alignment",
  wrap: "wrap",
} as const satisfies Record<
  keyof ParagraphFrameProperties,
  | "drop-cap"
  | "horizontal-alignment"
  | "horizontal-anchor"
  | "number"
  | "vertical-alignment"
  | "vertical-anchor"
  | "wrap"
>;
const FRAME_KEYS = new Set<string>(Object.keys(FRAME_VALIDATION));

const validateFrame = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, FRAME_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(FRAME_VALIDATION)) {
    const property = value[key];
    if (property === undefined) {
      continue;
    }
    if (validation === "number" && !isFiniteNumber(property)) {
      return INVALID;
    }
    if (validation === "drop-cap" && !isOneOf(property, ["none", "drop", "margin"])) {
      return INVALID;
    }
    if (validation === "horizontal-anchor" && !isOneOf(property, ["text", "margin", "page"])) {
      return INVALID;
    }
    if (validation === "vertical-anchor" && !isOneOf(property, ["text", "margin", "page"])) {
      return INVALID;
    }
    if (validation === "horizontal-alignment" && !isOneOf(property, FRAME_X_ALIGN_VALUES)) {
      return INVALID;
    }
    if (validation === "vertical-alignment" && !isOneOf(property, FRAME_Y_ALIGN_VALUES)) {
      return INVALID;
    }
    if (validation === "wrap" && !isOneOf(property, FRAME_WRAP_VALUES)) {
      return INVALID;
    }
  }
  return VALID;
};

const validateAuthoredProperty = (
  key: AuthoredParagraphPropertyKey,
  value: unknown,
): ValidationResult => {
  switch (key) {
    case "bidi":
    case "kinsoku":
    case "overflowPunctuation":
    case "snapToGrid":
    case "beforeAutospacing":
    case "afterAutospacing":
    case "hangingIndent":
    case "keepNext":
    case "keepLines":
    case "widowControl":
    case "pageBreakBefore":
    case "contextualSpacing":
    case "suppressLineNumbers":
    case "suppressAutoHyphens":
      return typeof value === "boolean" ? VALID : INVALID;
    case "spaceBefore":
    case "spaceAfter":
    case "lineSpacing":
    case "indentLeft":
    case "indentRight":
    case "indentFirstLine":
    case "outlineLevel":
      return isFiniteNumber(value) ? VALID : INVALID;
    case "alignment":
      return isOneOf(value, PARAGRAPH_ALIGNMENT_VALUES) ? VALID : INVALID;
    case "lineSpacingRule":
      return isOneOf(value, LINE_SPACING_RULE_VALUES) ? VALID : INVALID;
    case "spacingExplicit":
      return validateOptionalBooleanRecord(value, SPACING_EXPLICIT_KEY_LIST);
    case "borders":
      return validateBorders(value);
    case "shading":
      return validateShading(value);
    case "tabs":
      return validateTabs(value);
    case "numPr":
      return validateNumPr(value);
    case "styleId":
      return typeof value === "string" ? VALID : INVALID;
    case "frame":
      return validateFrame(value);
    default: {
      const exhaustive: never = key;
      return exhaustive;
    }
  }
};

const isAuthoredParagraphPropertyKey = (key: string): key is AuthoredParagraphPropertyKey => {
  const descriptor = Reflect.get(PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR, key) as
    | { owner?: unknown }
    | undefined;
  return descriptor?.owner === "pPr-base";
};

const validateAuthoredParagraphProperties = (value: unknown): ValidationResult => {
  if (!isRecord(value)) {
    return INVALID;
  }
  for (const [key, propertyValue] of Object.entries(value)) {
    if (!isAuthoredParagraphPropertyKey(key) || propertyValue === undefined) {
      return INVALID;
    }
    if (!validateAuthoredProperty(key, propertyValue).valid) {
      return INVALID;
    }
  }
  if (
    value["hangingIndent"] === true &&
    (!isFiniteNumber(value["indentFirstLine"]) || value["indentFirstLine"] >= 0)
  ) {
    return INVALID;
  }
  if (
    isFiniteNumber(value["indentFirstLine"]) &&
    value["indentFirstLine"] < 0 &&
    value["hangingIndent"] !== true
  ) {
    return INVALID;
  }
  return VALID;
};

export const readAuthoredParagraphProperties = (
  raw: unknown,
): ParagraphPropertySourceAttribute<AuthoredParagraphProperties> => {
  if (raw === null || raw === undefined) {
    return { status: "absent" };
  }
  if (!validateAuthoredParagraphProperties(raw).valid) {
    return { raw, status: "invalid" };
  }
  // SAFETY: validation established the descriptor-owned property shapes.
  return {
    status: "valid",
    value: cloneJsonValue(raw) as AuthoredParagraphProperties,
  };
};

const cloneJsonValue = (value: unknown): unknown => {
  if (Array.isArray(value)) {
    return value.map(cloneJsonValue);
  }
  if (isRecord(value)) {
    const clone: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value)) {
      clone[key] = cloneJsonValue(entry);
    }
    return clone;
  }
  return value;
};

const deepFreezeJsonValue = (value: unknown): unknown => {
  if (Array.isArray(value)) {
    for (const entry of value) {
      deepFreezeJsonValue(entry);
    }
    return Object.freeze(value);
  }
  if (isRecord(value)) {
    for (const entry of Object.values(value)) {
      deepFreezeJsonValue(entry);
    }
    return Object.freeze(value);
  }
  return value;
};

const cloneAuthoredParagraphProperties = (
  authoredPPr: AuthoredParagraphProperties,
): AuthoredParagraphProperties => {
  if (!validateAuthoredParagraphProperties(authoredPPr).valid) {
    panic("Invalid authored paragraph properties");
  }
  // SAFETY: the exhaustive validator above proves the clone contains only
  // descriptor-owned fields with their exact ParagraphFormatting value shapes.
  return deepFreezeJsonValue(cloneJsonValue(authoredPPr)) as AuthoredParagraphProperties;
};

type NumberingLevelIndentGeometryKey = keyof NumberingLevelIndentGeometry;
const NUMBERING_LEVEL_INDENT_GEOMETRY_VALIDATION = {
  indentLeft: "number",
  indentRight: "number",
  indentFirstLine: "number",
  hangingIndent: "boolean",
} as const satisfies Record<NumberingLevelIndentGeometryKey, "boolean" | "number">;
const NUMBERING_LEVEL_INDENT_GEOMETRY_KEYS = new Set<string>(
  Object.keys(NUMBERING_LEVEL_INDENT_GEOMETRY_VALIDATION),
);

const validateFirstLineIndent = (value: Record<string, unknown>): ValidationResult => {
  const firstLine = value["indentFirstLine"];
  const hanging = value["hangingIndent"];
  if (hanging === true && (!isFiniteNumber(firstLine) || firstLine >= 0)) {
    return INVALID;
  }
  return isFiniteNumber(firstLine) && firstLine < 0 && hanging !== true ? INVALID : VALID;
};

const validateNumberingLevelIndentGeometry = (
  value: unknown,
  requireNonEmpty: boolean,
): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, NUMBERING_LEVEL_INDENT_GEOMETRY_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(NUMBERING_LEVEL_INDENT_GEOMETRY_VALIDATION)) {
    const property = value[key];
    if (property === undefined) {
      continue;
    }
    if (
      (validation === "number" && !isFiniteNumber(property)) ||
      (validation === "boolean" && typeof property !== "boolean")
    ) {
      return INVALID;
    }
  }
  if (
    requireNonEmpty &&
    value["indentLeft"] === undefined &&
    value["indentRight"] === undefined &&
    value["indentFirstLine"] === undefined
  ) {
    return INVALID;
  }
  return validateFirstLineIndent(value);
};

type LatentNumberingLevelIndent = Extract<NumberingLevelIndentProvenance, { type: "latent" }>;
type OwnedNumberingLevelIndent = Extract<NumberingLevelIndentProvenance, { type: "owned" }>;
const LATENT_NUMBERING_LEVEL_INDENT_VALIDATION = {
  type: "latent",
  numId: "identity",
  ilvl: "identity",
  baseline: "geometry",
} as const satisfies Record<keyof LatentNumberingLevelIndent, string>;
const OWNED_NUMBERING_LEVEL_INDENT_VALIDATION = {
  type: "owned",
  numId: "identity",
  ilvl: "identity",
  baseline: "geometry",
  owned: "geometry",
} as const satisfies Record<keyof OwnedNumberingLevelIndent, string>;
const LATENT_NUMBERING_LEVEL_INDENT_KEYS = new Set<string>(
  Object.keys(LATENT_NUMBERING_LEVEL_INDENT_VALIDATION),
);
const OWNED_NUMBERING_LEVEL_INDENT_KEYS = new Set<string>(
  Object.keys(OWNED_NUMBERING_LEVEL_INDENT_VALIDATION),
);

const validateNumberingLevelIndent = (value: unknown): ValidationResult => {
  if (!isRecord(value) || (value["type"] !== "latent" && value["type"] !== "owned")) {
    return INVALID;
  }
  const keys =
    value["type"] === "latent"
      ? LATENT_NUMBERING_LEVEL_INDENT_KEYS
      : OWNED_NUMBERING_LEVEL_INDENT_KEYS;
  if (
    !hasOnlyKeys(value, keys) ||
    !isNonNegativeSafeInteger(value["numId"]) ||
    value["numId"] === 0 ||
    !isNonNegativeSafeInteger(value["ilvl"]) ||
    !validateNumberingLevelIndentGeometry(value["baseline"], true).valid
  ) {
    return INVALID;
  }
  if (value["type"] === "latent") {
    return VALID;
  }
  if (!validateNumberingLevelIndentGeometry(value["owned"], true).valid) {
    return INVALID;
  }
  const baseline = value["baseline"] as NonEmptyNumberingLevelIndentGeometry;
  const owned = value["owned"] as NonEmptyNumberingLevelIndentGeometry;
  for (const key of Object.keys(
    NUMBERING_LEVEL_INDENT_GEOMETRY_VALIDATION,
  ) as NumberingLevelIndentGeometryKey[]) {
    if (owned[key] !== undefined && owned[key] !== baseline[key]) {
      return INVALID;
    }
  }
  return VALID;
};

const PARAGRAPH_MARK_EFFECTIVE_VALIDATION = {
  defaultTextFormatting: "text-formatting",
  runInWithNext: "boolean",
} as const satisfies Record<keyof ParagraphMarkEffectiveProperties, "boolean" | "text-formatting">;
const PARAGRAPH_MARK_EFFECTIVE_KEYS = new Set<string>(
  Object.keys(PARAGRAPH_MARK_EFFECTIVE_VALIDATION),
);

const validateParagraphMarkEffectiveProperties = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, PARAGRAPH_MARK_EFFECTIVE_KEYS)) {
    return INVALID;
  }
  for (const [key, validation] of Object.entries(PARAGRAPH_MARK_EFFECTIVE_VALIDATION)) {
    const property = value[key];
    if (property === undefined) {
      continue;
    }
    if (
      (validation === "boolean" && typeof property !== "boolean") ||
      (validation === "text-formatting" && !validateTextFormatting(property).valid)
    ) {
      return INVALID;
    }
  }
  return VALID;
};

const PARAGRAPH_MARK_CONTEXT_VALIDATION = {
  authored: "paragraph-mark",
  effective: "effective",
} as const satisfies Record<keyof ParagraphMarkProjectionContext, "effective" | "paragraph-mark">;
const PARAGRAPH_MARK_CONTEXT_KEYS = new Set<string>(Object.keys(PARAGRAPH_MARK_CONTEXT_VALIDATION));

const validateParagraphMarkContext = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, PARAGRAPH_MARK_CONTEXT_KEYS)) {
    return INVALID;
  }
  return readParagraphMarkProperties(value["authored"]).status === "valid" &&
    validateParagraphMarkEffectiveProperties(value["effective"]).valid
    ? VALID
    : INVALID;
};

const PARAGRAPH_SPACING_INHERITANCE_VALIDATION = {
  before: "spacing-source",
  after: "spacing-source",
} as const satisfies Record<keyof ParagraphSpacingInheritance, "spacing-source">;
const PARAGRAPH_SPACING_INHERITANCE_KEYS = new Set<string>(
  Object.keys(PARAGRAPH_SPACING_INHERITANCE_VALIDATION),
);
const PARAGRAPH_SPACING_INHERITANCE_VALUES = Object.freeze(
  Object.values(PARAGRAPH_SPACING_INHERITANCE_SOURCE),
);

const validateParagraphSpacingInheritance = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, PARAGRAPH_SPACING_INHERITANCE_KEYS)) {
    return INVALID;
  }
  return Object.values(value).every((entry) => isOneOf(entry, PARAGRAPH_SPACING_INHERITANCE_VALUES))
    ? VALID
    : INVALID;
};

const PARAGRAPH_PROPERTY_CONTEXT_VALIDATION = {
  inheritedPPr: "authored-ppr",
  numberingLevelIndent: "numbering-indent",
  numPrFromStyle: "num-pr",
  paragraphMark: "paragraph-mark",
  spacingInheritance: "spacing-inheritance",
} as const satisfies Record<keyof SerializedParagraphPropertyProjectionContext, string>;
const PARAGRAPH_PROPERTY_CONTEXT_KEYS = new Set<string>(
  Object.keys(PARAGRAPH_PROPERTY_CONTEXT_VALIDATION),
);

const validateParagraphPropertyProjectionContext = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, PARAGRAPH_PROPERTY_CONTEXT_KEYS)) {
    return INVALID;
  }
  if (
    readAuthoredParagraphProperties(value["inheritedPPr"]).status !== "valid" ||
    !validateParagraphMarkContext(value["paragraphMark"]).valid ||
    !validateParagraphSpacingInheritance(value["spacingInheritance"]).valid
  ) {
    return INVALID;
  }
  const numberingLevelIndent = value["numberingLevelIndent"];
  if (numberingLevelIndent !== null && !validateNumberingLevelIndent(numberingLevelIndent).valid) {
    return INVALID;
  }
  const numPrFromStyle = value["numPrFromStyle"];
  if (numPrFromStyle !== null && !validateNumPr(numPrFromStyle).valid) {
    return INVALID;
  }
  const inheritedPPr = value["inheritedPPr"] as AuthoredParagraphProperties;
  const spacingInheritance = value["spacingInheritance"] as ParagraphSpacingInheritance;
  if (
    (spacingInheritance.before !== undefined && inheritedPPr.spaceBefore === undefined) ||
    (spacingInheritance.after !== undefined && inheritedPPr.spaceAfter === undefined)
  ) {
    return INVALID;
  }
  return VALID;
};

const cloneParagraphPropertyProjectionContext = (
  context: SerializedParagraphPropertyProjectionContext,
): ParagraphPropertyProjectionContext => {
  if (!validateParagraphPropertyProjectionContext(context).valid) {
    panic("Invalid paragraph-property projection context");
  }
  // SAFETY: the total context and nested validators establish the complete shape.
  return deepFreezeJsonValue(cloneJsonValue(context)) as ParagraphPropertyProjectionContext;
};

const effectiveNumberingIdentity = (
  authoredPPr: DeepReadonly<AuthoredParagraphProperties>,
  context: ParagraphPropertyProjectionContext,
): ParagraphFormatting["numPr"] | undefined => {
  const inherited = context.inheritedPPr.numPr;
  const authored = authoredPPr.numPr;
  if (inherited === undefined && authored === undefined) {
    return undefined;
  }
  return { ...inherited, ...authored };
};

const paragraphPropertyStatePayloadIsValid = (
  authoredPPr: AuthoredParagraphProperties,
  context: ParagraphPropertyProjectionContext,
): boolean => {
  const provenance = context.numberingLevelIndent;
  if (provenance === null) {
    return true;
  }
  const numbering = effectiveNumberingIdentity(authoredPPr, context);
  return !(
    numbering?.numId === undefined ||
    numbering.numId === 0 ||
    numbering.numId !== provenance.numId ||
    (numbering.ilvl ?? 0) !== provenance.ilvl
  );
};

const assertParagraphPropertyStatePayload = (
  authoredPPr: AuthoredParagraphProperties,
  context: ParagraphPropertyProjectionContext,
): void => {
  if (!paragraphPropertyStatePayloadIsValid(authoredPPr, context)) {
    panic("Numbering indentation provenance must match effective paragraph numbering");
  }
};

const trustedParagraphPropertyPayload = (
  authoredPPr: AuthoredParagraphProperties,
  context: SerializedParagraphPropertyProjectionContext,
): Pick<TrustedParagraphPropertyState, "authoredPPr" | "context"> => {
  const trustedAuthored = cloneAuthoredParagraphProperties(authoredPPr);
  const trustedContext = cloneParagraphPropertyProjectionContext(context);
  assertParagraphPropertyStatePayload(trustedAuthored, trustedContext);
  return { authoredPPr: trustedAuthored, context: trustedContext };
};

type UntrustedStateBranch =
  | {
      type: "imported";
      token: string;
      authoredPPr: AuthoredParagraphProperties;
      context: ParagraphPropertyProjectionContext;
    }
  | {
      type: "editor-created";
      authoredPPr: AuthoredParagraphProperties;
      context: ParagraphPropertyProjectionContext;
    }
  | {
      type: "transient-template";
      handle: ParagraphPropertyTransientTemplateHandle;
      authoredPPr: AuthoredParagraphProperties;
      context: ParagraphPropertyProjectionContext;
    };

const trustParagraphPropertyState = <Branch extends UntrustedStateBranch>(
  state: Branch,
): Branch & TrustedParagraphPropertyState => {
  Object.defineProperty(state, TRUSTED_PARAGRAPH_PROPERTY_STATE, {
    configurable: false,
    enumerable: false,
    value: true,
    writable: false,
  });
  if (state.type === "transient-template") {
    Object.defineProperty(state, "toJSON", {
      configurable: false,
      enumerable: false,
      value: () => panic("Transient paragraph-property state cannot cross a persistence boundary"),
      writable: false,
    });
  }
  Object.freeze(state);
  // SAFETY: this function alone installs the private runtime brand after all
  // payloads have passed validation and deep freezing.
  return state as Branch & TrustedParagraphPropertyState;
};

const stateKeysAreValid = (value: Record<string, unknown>, type: string): boolean => {
  switch (type) {
    case "imported":
      return hasOnlyKeys(value, IMPORTED_STATE_KEYS);
    case "editor-created":
      return hasOnlyKeys(value, EDITOR_STATE_KEYS);
    default:
      return false;
  }
};

const IMPORTED_STATE_KEYS = new Set<string>(["type", "token", "authoredPPr", "context"]);
const EDITOR_STATE_KEYS = new Set<string>(["type", "authoredPPr", "context"]);

type CreateImportedParagraphPropertyStateOptions = {
  token: string;
  authoredPPr: AuthoredParagraphProperties;
  context: SerializedParagraphPropertyProjectionContext;
};

export const createImportedParagraphPropertyState = ({
  token,
  authoredPPr,
  context,
}: CreateImportedParagraphPropertyStateOptions): PersistableParagraphPropertyState =>
  trustParagraphPropertyState({
    type: "imported",
    token,
    ...trustedParagraphPropertyPayload(authoredPPr, context),
  });

type CreateEditorParagraphPropertyStateOptions = {
  authoredPPr?: AuthoredParagraphProperties;
  context?: SerializedParagraphPropertyProjectionContext;
};

export const createEditorParagraphPropertyState = ({
  authoredPPr = {},
  context = EMPTY_PARAGRAPH_PROPERTY_CONTEXT,
}: CreateEditorParagraphPropertyStateOptions = {}): PersistableParagraphPropertyState =>
  trustParagraphPropertyState({
    type: "editor-created",
    ...trustedParagraphPropertyPayload(authoredPPr, context),
  });

/** Validate and freeze the model-level authored selector at the PM trust boundary. */
export const authoredParagraphPropertiesFromFormatting = (
  formatting: Parameters<typeof selectAuthoredParagraphProperties>[0],
): AuthoredParagraphProperties =>
  cloneAuthoredParagraphProperties(selectAuthoredParagraphProperties(formatting));

type CreateTransientTemplateParagraphPropertyStateOptions = {
  handle: ParagraphPropertyTransientTemplateHandle;
  authoredPPr: AuthoredParagraphProperties;
  context: SerializedParagraphPropertyProjectionContext;
};

export const createTransientTemplateParagraphPropertyState = ({
  handle,
  authoredPPr,
  context,
}: CreateTransientTemplateParagraphPropertyStateOptions): ParagraphPropertyState =>
  trustParagraphPropertyState({
    type: "transient-template",
    handle,
    ...trustedParagraphPropertyPayload(authoredPPr, context),
  });

const isTrustedParagraphPropertyState = (raw: unknown): raw is ParagraphPropertyState =>
  isRecord(raw) && Reflect.get(raw, TRUSTED_PARAGRAPH_PROPERTY_STATE) === true;

export const readParagraphPropertyState = (
  raw: unknown,
): ParagraphPropertySourceAttribute<ParagraphPropertyState> => {
  if (raw === null || raw === undefined) {
    return { status: "absent" };
  }
  if (isTrustedParagraphPropertyState(raw)) {
    return { status: "valid", value: raw };
  }
  assertParagraphPropertyStateBudget(raw);
  if (!isRecord(raw) || typeof raw["type"] !== "string") {
    return { raw, status: "invalid" };
  }
  const type = raw["type"];
  if (
    !stateKeysAreValid(raw, type) ||
    !validateAuthoredParagraphProperties(raw["authoredPPr"]).valid ||
    !validateParagraphPropertyProjectionContext(raw["context"]).valid
  ) {
    return { raw, status: "invalid" };
  }
  // SAFETY: the validation above established the complete descriptor-owned shape.
  const authoredPPr = cloneAuthoredParagraphProperties(
    raw["authoredPPr"] as AuthoredParagraphProperties,
  );
  const context = cloneParagraphPropertyProjectionContext(
    raw["context"] as SerializedParagraphPropertyProjectionContext,
  );
  if (!paragraphPropertyStatePayloadIsValid(authoredPPr, context)) {
    return { raw, status: "invalid" };
  }
  switch (type) {
    case "imported": {
      const token = raw["token"];
      return typeof token === "string" && token.length > 0
        ? {
            status: "valid",
            value: trustParagraphPropertyState({ type, token, authoredPPr, context }),
          }
        : { raw, status: "invalid" };
    }
    case "editor-created":
      return {
        status: "valid",
        value: trustParagraphPropertyState({ type, authoredPPr, context }),
      };
    default:
      return { raw, status: "invalid" };
  }
};

/** Reify state at a PM/Yjs persistence boundary, where template handles are invalid. */
export const readPersistableParagraphPropertyState = (
  raw: unknown,
): ParagraphPropertySourceAttribute<PersistableParagraphPropertyState> => {
  const result = readParagraphPropertyState(raw);
  if (result.status !== "valid") {
    return result;
  }
  if (result.value.type === "transient-template") {
    return { raw, status: "invalid" };
  }
  return result;
};

export const expectParagraphPropertyState = (raw: unknown): ParagraphPropertyState => {
  const result = readParagraphPropertyState(raw);
  if (result.status !== "valid") {
    panic("Paragraphs require a valid paragraph-property state");
  }
  return result.value;
};

/** Serialize only state safe to store in Yjs or another persistent document. */
export const serializePersistableParagraphPropertyState = (
  state: ParagraphPropertyState,
): SerializedPersistableParagraphPropertyState => {
  switch (state.type) {
    case "imported":
      return {
        type: "imported",
        token: state.token,
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
        context: cloneJsonValue(state.context) as SerializedParagraphPropertyProjectionContext,
      };
    case "editor-created":
      return {
        type: "editor-created",
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
        context: cloneJsonValue(state.context) as SerializedParagraphPropertyProjectionContext,
      };
    case "transient-template":
      return panic("Transient paragraph-property state cannot cross a persistence boundary");
    default: {
      const exhaustive: never = state;
      return exhaustive;
    }
  }
};

/** PM attrs carry wire-safe state, except a conversion-scoped nominal transient capsule. */
export type ParagraphPropertyStateAttribute =
  | SerializedParagraphPropertyState
  | Extract<ParagraphPropertyState, { type: "transient-template" }>;

export const paragraphPropertyStateAttribute = (
  state: ParagraphPropertyState,
): ParagraphPropertyStateAttribute =>
  state.type === "transient-template" ? state : serializePersistableParagraphPropertyState(state);

export const transitionParagraphPropertyState = (
  state: ParagraphPropertyState,
  transition: ParagraphPropertyStateTransition,
): ParagraphPropertyState => {
  const authoredPPr = applyAuthoredTransition(state.authoredPPr, transition.authored);
  const context = applyContextTransition(state.context, transition.context);
  assertParagraphPropertyStatePayload(authoredPPr, context);
  switch (state.type) {
    case "imported":
      return trustParagraphPropertyState({
        type: "imported",
        token: state.token,
        authoredPPr,
        context,
      });
    case "editor-created":
      return trustParagraphPropertyState({ type: "editor-created", authoredPPr, context });
    case "transient-template":
      return trustParagraphPropertyState({
        type: "transient-template",
        handle: state.handle,
        authoredPPr,
        context,
      });
    default: {
      const exhaustive: never = state;
      return exhaustive;
    }
  }
};

const applyAuthoredTransition = (
  authoredPPr: DeepReadonly<AuthoredParagraphProperties>,
  transition: ParagraphPropertyAuthoredTransition,
): AuthoredParagraphProperties => {
  switch (transition.type) {
    case "preserve":
      return cloneAuthoredParagraphProperties(authoredPPr);
    case "mutate":
      return applyParagraphPropertyMutations(authoredPPr, transition.mutations);
    case "replace":
      return cloneAuthoredParagraphProperties(transition.authoredPPr);
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }
};

const applyParagraphMarkMutations = (
  mark: ParagraphMarkProperties,
  mutations: readonly ParagraphMarkPropertyMutation[],
): ParagraphMarkProperties => {
  const next = cloneJsonValue(mark) as ParagraphMarkProperties;
  const seen = new Set<ParagraphMarkPropertyMutation["key"]>();
  for (const { key, mutation } of mutations) {
    if (seen.has(key)) {
      panic(`Paragraph-mark mutation batch contains duplicate key: ${key}`);
    }
    seen.add(key);
    if (mutation.type === "remove") {
      Reflect.deleteProperty(next, key);
    } else {
      Reflect.set(next, key, cloneJsonValue(mutation.value));
    }
  }
  const validated = readParagraphMarkProperties(next);
  if (validated.status !== "valid") {
    panic("Invalid paragraph-mark mutation result");
  }
  return validated.value;
};

const applyParagraphMarkEffectiveMutations = (
  mark: ParagraphMarkEffectiveProperties,
  mutations: readonly ParagraphMarkEffectivePropertyMutation[],
): ParagraphMarkEffectiveProperties => {
  const next = cloneJsonValue(mark) as ParagraphMarkEffectiveProperties;
  const seen = new Set<ParagraphMarkEffectivePropertyMutation["key"]>();
  for (const { key, mutation } of mutations) {
    if (seen.has(key)) {
      panic(`Effective paragraph-mark mutation batch contains duplicate key: ${key}`);
    }
    seen.add(key);
    if (mutation.type === "remove") {
      Reflect.deleteProperty(next, key);
    } else {
      Reflect.set(next, key, cloneJsonValue(mutation.value));
    }
  }
  if (!validateParagraphMarkEffectiveProperties(next).valid) {
    panic("Invalid effective paragraph-mark mutation result");
  }
  return next;
};

const applyContextTransition = (
  context: ParagraphPropertyProjectionContext,
  transition: ParagraphPropertyContextTransition,
): ParagraphPropertyProjectionContext => {
  switch (transition.type) {
    case "preserve":
      return cloneParagraphPropertyProjectionContext(context);
    case "replace":
      return cloneParagraphPropertyProjectionContext(transition.context);
    case "mutate-paragraph-mark":
      return cloneParagraphPropertyProjectionContext({
        ...context,
        paragraphMark: {
          authored: applyParagraphMarkMutations(
            context.paragraphMark.authored,
            transition.authored,
          ),
          effective: applyParagraphMarkEffectiveMutations(
            context.paragraphMark.effective,
            transition.effective,
          ),
        },
      });
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }
};

export const applyParagraphPropertyMutations = (
  authoredPPr: DeepReadonly<AuthoredParagraphProperties>,
  mutations: readonly ParagraphPropertyMutation[],
): AuthoredParagraphProperties => {
  if (!validateAuthoredParagraphProperties(authoredPPr).valid) {
    panic("Invalid authored paragraph properties");
  }
  // SAFETY: the source passed the exhaustive authored-property validator.
  const next = cloneJsonValue(authoredPPr) as AuthoredParagraphProperties;
  const seen = new Set<AuthoredParagraphPropertyKey>();
  for (const { key, mutation } of mutations) {
    if (!isAuthoredParagraphPropertyKey(key)) {
      panic(`Cannot mutate non-authored paragraph property: ${key}`);
    }
    if (seen.has(key)) {
      panic(`Paragraph-property mutation batch contains duplicate key: ${key}`);
    }
    seen.add(key);
    switch (mutation.type) {
      case "remove":
        Reflect.deleteProperty(next, key);
        break;
      case "set":
        Reflect.set(next, key, cloneJsonValue(mutation.value));
        break;
      default: {
        const exhaustive: never = mutation;
        return exhaustive;
      }
    }
  }
  return cloneAuthoredParagraphProperties(next);
};

export const splitParagraphPropertyState = (
  state: ParagraphPropertyState,
  transition: ParagraphPropertySplitTransition,
): {
  type: ParagraphPropertySplitTransition["type"];
  left: PersistableParagraphPropertyState;
  right: ParagraphPropertyState;
} => {
  switch (transition.type) {
    case "split-left-created-right-retains":
      return {
        type: transition.type,
        left: createEditorParagraphPropertyState({
          authoredPPr: state.authoredPPr,
          context: state.context,
        }),
        right: state,
      };
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }
};

export const joinParagraphPropertyStates = (
  left: ParagraphPropertyState,
  right: ParagraphPropertyState,
  transition: ParagraphPropertyJoinTransition,
): ParagraphPropertyState => {
  switch (transition.type) {
    case "join-left-paragraph-mark-retains":
      return left;
    case "join-right-paragraph-mark-retains":
      return right;
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }
};

export const EMPTY_PARAGRAPH_PROPERTY_CONTEXT = Object.freeze({
  inheritedPPr: Object.freeze({}),
  numberingLevelIndent: null,
  numPrFromStyle: null,
  paragraphMark: Object.freeze({
    authored: Object.freeze({}),
    effective: Object.freeze({}),
  }),
  spacingInheritance: Object.freeze({}),
}) satisfies SerializedParagraphPropertyProjectionContext;

export const EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE = Object.freeze({
  type: "editor-created",
  authoredPPr: Object.freeze({}),
  context: EMPTY_PARAGRAPH_PROPERTY_CONTEXT,
}) satisfies SerializedPersistableParagraphPropertyState;
