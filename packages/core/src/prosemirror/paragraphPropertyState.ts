import { panic } from "better-result";

import {
  ParagraphPropertySourceToken,
  ParagraphPropertyTransientTemplateHandle,
  type ParagraphPropertySourceAttribute,
} from "../docx/paragraphPropertySourceIdentity";
import {
  FRAME_WRAP_VALUES,
  FRAME_X_ALIGN_VALUES,
  FRAME_Y_ALIGN_VALUES,
  LINE_SPACING_RULE_VALUES,
  PARAGRAPH_ALIGNMENT_VALUES,
  SHADING_PATTERN_VALUES,
  TAB_LEADER_VALUES,
  TAB_STOP_ALIGNMENT_VALUES,
  THEME_COLOR_SLOT_VALUES,
} from "../types/documentEnumValues";
import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  selectAuthoredParagraphProperties,
  type AuthoredParagraphProperties,
  type AuthoredParagraphPropertyKey,
} from "../docx/paragraphPropertyDescriptor";

export type SerializedPersistableParagraphPropertyState =
  | {
      type: "imported";
      token: string;
      authoredPPr: AuthoredParagraphProperties;
    }
  | {
      type: "editor-created";
      authoredPPr: AuthoredParagraphProperties;
    };

/** Internal PM-only state; transient handles are forbidden at persistence boundaries. */
export type SerializedParagraphPropertyState =
  | SerializedPersistableParagraphPropertyState
  | {
      type: "transient-template";
      handle: string;
      authoredPPr: AuthoredParagraphProperties;
    };

const TRUSTED_PARAGRAPH_PROPERTY_STATE: unique symbol = Symbol(
  "trusted-paragraph-property-state",
);

type TrustedParagraphPropertyState = {
  readonly [TRUSTED_PARAGRAPH_PROPERTY_STATE]: true;
};

export type PersistableParagraphPropertyState = (
  | {
      readonly type: "imported";
      readonly token: ParagraphPropertySourceToken;
      readonly authoredPPr: Readonly<AuthoredParagraphProperties>;
    }
  | {
      readonly type: "editor-created";
      readonly authoredPPr: Readonly<AuthoredParagraphProperties>;
    }
  ) & TrustedParagraphPropertyState;

/** Trusted state used after the serialized PM/Yjs boundary has been validated. */
export type ParagraphPropertyState =
  | PersistableParagraphPropertyState
  | ({
      readonly type: "transient-template";
      readonly handle: ParagraphPropertyTransientTemplateHandle;
      readonly authoredPPr: Readonly<AuthoredParagraphProperties>;
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

export type ParagraphPropertyStateTransition =
  | { type: "mutate-authored"; mutations: readonly ParagraphPropertyMutation[] }
  | { type: "replace-authored"; authoredPPr: AuthoredParagraphProperties }
  | { type: "editor-copy" };

export type ParagraphPropertySplitTransition = {
  type: "split-left-created-right-retains";
};

export type ParagraphPropertyJoinTransition =
  | { type: "join-left-paragraph-mark-retains" }
  | { type: "join-right-paragraph-mark-retains" };

type ValidationResult = { valid: true } | { valid: false };

const INVALID: ValidationResult = { valid: false };
const VALID: ValidationResult = { valid: true };

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

const COLOR_VALUE_KEYS = new Set(["rgb", "auto", "themeColor", "themeTint", "themeShade"]);

const validateColorValue = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, COLOR_VALUE_KEYS)) {
    return INVALID;
  }
  if (value["rgb"] !== undefined && typeof value["rgb"] !== "string") {
    return INVALID;
  }
  if (value["auto"] !== undefined && typeof value["auto"] !== "boolean") {
    return INVALID;
  }
  if (
    value["themeColor"] !== undefined &&
    !isOneOf(value["themeColor"], THEME_COLOR_SLOT_VALUES)
  ) {
    return INVALID;
  }
  if (value["themeTint"] !== undefined && typeof value["themeTint"] !== "string") {
    return INVALID;
  }
  if (value["themeShade"] !== undefined && typeof value["themeShade"] !== "string") {
    return INVALID;
  }
  return VALID;
};

const BORDER_KEYS = new Set(["style", "size", "space", "color", "shadow", "frame"]);

const validateBorder = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, BORDER_KEYS)) {
    return INVALID;
  }
  if (value["style"] !== undefined && typeof value["style"] !== "string") {
    return INVALID;
  }
  if (value["size"] !== undefined && !isFiniteNumber(value["size"])) {
    return INVALID;
  }
  if (value["space"] !== undefined && !isFiniteNumber(value["space"])) {
    return INVALID;
  }
  if (value["shadow"] !== undefined && typeof value["shadow"] !== "boolean") {
    return INVALID;
  }
  if (value["frame"] !== undefined && typeof value["frame"] !== "boolean") {
    return INVALID;
  }
  if (value["color"] !== undefined && !validateColorValue(value["color"]).valid) {
    return INVALID;
  }
  return VALID;
};

const PARAGRAPH_BORDER_KEYS = new Set(["top", "bottom", "left", "right", "between", "bar"]);

const validateBorders = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, PARAGRAPH_BORDER_KEYS)) {
    return INVALID;
  }
  return Object.values(value).every((entry) => validateBorder(entry).valid) ? VALID : INVALID;
};

const SHADING_KEYS = new Set(["color", "fill", "pattern"]);

const validateShading = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, SHADING_KEYS)) {
    return INVALID;
  }
  if (value["color"] !== undefined && !validateColorValue(value["color"]).valid) {
    return INVALID;
  }
  if (value["fill"] !== undefined && !validateColorValue(value["fill"]).valid) {
    return INVALID;
  }
  if (value["pattern"] !== undefined && !isOneOf(value["pattern"], SHADING_PATTERN_VALUES)) {
    return INVALID;
  }
  return VALID;
};

const TAB_KEYS = new Set(["position", "alignment", "leader"]);

const validateTabs = (value: unknown): ValidationResult => {
  if (!Array.isArray(value)) {
    return INVALID;
  }
  for (const tab of value) {
    if (!isRecord(tab) || !hasOnlyKeys(tab, TAB_KEYS)) {
      return INVALID;
    }
    if (!isFiniteNumber(tab["position"])) {
      return INVALID;
    }
    if (!isOneOf(tab["alignment"], TAB_STOP_ALIGNMENT_VALUES)) {
      return INVALID;
    }
    if (tab["leader"] !== undefined && !isOneOf(tab["leader"], TAB_LEADER_VALUES)) {
      return INVALID;
    }
  }
  return VALID;
};

const NUM_PR_KEYS = new Set(["numId", "ilvl"]);

const validateNumPr = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, NUM_PR_KEYS)) {
    return INVALID;
  }
  if (value["numId"] !== undefined && !isNonNegativeSafeInteger(value["numId"])) {
    return INVALID;
  }
  if (value["ilvl"] !== undefined && !isNonNegativeSafeInteger(value["ilvl"])) {
    return INVALID;
  }
  return VALID;
};

const FRAME_KEYS = new Set([
  "dropCap",
  "lines",
  "width",
  "height",
  "hSpace",
  "vSpace",
  "hAnchor",
  "vAnchor",
  "x",
  "y",
  "xAlign",
  "yAlign",
  "wrap",
]);

const validateFrame = (value: unknown): ValidationResult => {
  if (!isRecord(value) || !hasOnlyKeys(value, FRAME_KEYS)) {
    return INVALID;
  }
  if (value["dropCap"] !== undefined && !isOneOf(value["dropCap"], ["none", "drop", "margin"])) {
    return INVALID;
  }
  if (value["hAnchor"] !== undefined && !isOneOf(value["hAnchor"], ["text", "margin", "page"])) {
    return INVALID;
  }
  if (value["vAnchor"] !== undefined && !isOneOf(value["vAnchor"], ["text", "margin", "page"])) {
    return INVALID;
  }
  if (value["xAlign"] !== undefined && !isOneOf(value["xAlign"], FRAME_X_ALIGN_VALUES)) {
    return INVALID;
  }
  if (value["yAlign"] !== undefined && !isOneOf(value["yAlign"], FRAME_Y_ALIGN_VALUES)) {
    return INVALID;
  }
  if (value["wrap"] !== undefined && !isOneOf(value["wrap"], FRAME_WRAP_VALUES)) {
    return INVALID;
  }
  for (const key of ["lines", "width", "height", "hSpace", "vSpace", "x", "y"] as const) {
    if (value[key] !== undefined && !isFiniteNumber(value[key])) {
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
      return validateOptionalBooleanRecord(value, ["before", "after"]);
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

type UntrustedStateBranch =
  | {
      type: "imported";
      token: ParagraphPropertySourceToken;
      authoredPPr: AuthoredParagraphProperties;
    }
  | {
      type: "editor-created";
      authoredPPr: AuthoredParagraphProperties;
    }
  | {
      type: "transient-template";
      handle: ParagraphPropertyTransientTemplateHandle;
      authoredPPr: AuthoredParagraphProperties;
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
  Object.freeze(state);
  // SAFETY: this function alone installs the private runtime brand after all
  // payloads have passed validation and deep freezing.
  return state as Branch & TrustedParagraphPropertyState;
};

const stateKeysAreValid = (value: Record<string, unknown>, type: string): boolean => {
  switch (type) {
    case "imported":
      return hasOnlyKeys(value, new Set(["type", "token", "authoredPPr"]));
    case "editor-created":
      return hasOnlyKeys(value, new Set(["type", "authoredPPr"]));
    case "transient-template":
      return hasOnlyKeys(value, new Set(["type", "handle", "authoredPPr"]));
    default:
      return false;
  }
};

export const createImportedParagraphPropertyState = (
  token: ParagraphPropertySourceToken,
  authoredPPr: AuthoredParagraphProperties,
): PersistableParagraphPropertyState =>
  trustParagraphPropertyState({
    type: "imported",
    token,
    authoredPPr: cloneAuthoredParagraphProperties(authoredPPr),
  });

export const createEditorParagraphPropertyState = (
  authoredPPr: AuthoredParagraphProperties = {},
): PersistableParagraphPropertyState =>
  trustParagraphPropertyState({
    type: "editor-created",
    authoredPPr: cloneAuthoredParagraphProperties(authoredPPr),
  });

/** Validate and freeze the model-level authored selector at the PM trust boundary. */
export const authoredParagraphPropertiesFromFormatting = (
  formatting: Parameters<typeof selectAuthoredParagraphProperties>[0],
): AuthoredParagraphProperties =>
  cloneAuthoredParagraphProperties(selectAuthoredParagraphProperties(formatting));

export const createTransientTemplateParagraphPropertyState = (
  handle: ParagraphPropertyTransientTemplateHandle,
  authoredPPr: AuthoredParagraphProperties,
): ParagraphPropertyState =>
  trustParagraphPropertyState({
    type: "transient-template",
    handle,
    authoredPPr: cloneAuthoredParagraphProperties(authoredPPr),
  });

export const readParagraphPropertyState = (
  raw: unknown,
): ParagraphPropertySourceAttribute<ParagraphPropertyState> => {
  if (raw === null || raw === undefined) {
    return { status: "absent" };
  }
  if (!isRecord(raw) || typeof raw["type"] !== "string") {
    return { raw, status: "invalid" };
  }
  const type = raw["type"];
  if (!stateKeysAreValid(raw, type) || !validateAuthoredParagraphProperties(raw["authoredPPr"]).valid) {
    return { raw, status: "invalid" };
  }
  // SAFETY: the validation above established the complete descriptor-owned shape.
  const authoredPPr = cloneAuthoredParagraphProperties(
    raw["authoredPPr"] as AuthoredParagraphProperties,
  );
  switch (type) {
    case "imported": {
      const token = ParagraphPropertySourceToken.read(raw["token"]);
      return token.status === "valid"
        ? {
            status: "valid",
            value: trustParagraphPropertyState({ type, token: token.value, authoredPPr }),
          }
        : { raw, status: "invalid" };
    }
    case "editor-created":
      return {
        status: "valid",
        value: trustParagraphPropertyState({ type, authoredPPr }),
      };
    case "transient-template": {
      const handle = ParagraphPropertyTransientTemplateHandle.read(raw["handle"]);
      return handle.status === "valid"
        ? {
            status: "valid",
            value: trustParagraphPropertyState({ type, handle: handle.value, authoredPPr }),
          }
        : { raw, status: "invalid" };
    }
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

export const serializeParagraphPropertyState = (
  state: ParagraphPropertyState,
): SerializedParagraphPropertyState => {
  switch (state.type) {
    case "imported":
      return {
        type: "imported",
        token: state.token.serialized,
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
      };
    case "editor-created":
      return {
        type: "editor-created",
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
      };
    case "transient-template":
      return {
        type: "transient-template",
        handle: state.handle.serialized,
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
      };
    default: {
      const exhaustive: never = state;
      return exhaustive;
    }
  }
};

/** Serialize only state safe to store in Yjs or another persistent document. */
export const serializePersistableParagraphPropertyState = (
  state: ParagraphPropertyState,
): SerializedPersistableParagraphPropertyState => {
  switch (state.type) {
    case "imported":
      return {
        type: "imported",
        token: state.token.serialized,
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
      };
    case "editor-created":
      return {
        type: "editor-created",
        authoredPPr: cloneAuthoredParagraphProperties(state.authoredPPr),
      };
    case "transient-template":
      return panic("Transient paragraph-property state cannot cross a persistence boundary");
    default: {
      const exhaustive: never = state;
      return exhaustive;
    }
  }
};

export const transitionParagraphPropertyState = (
  state: ParagraphPropertyState,
  transition: ParagraphPropertyStateTransition,
): ParagraphPropertyState => {
  if (transition.type === "editor-copy") {
    return createEditorParagraphPropertyState(state.authoredPPr);
  }
  const authoredPPr =
    transition.type === "replace-authored"
      ? cloneAuthoredParagraphProperties(transition.authoredPPr)
      : applyParagraphPropertyMutations(state.authoredPPr, transition.mutations);
  switch (state.type) {
    case "imported":
      return trustParagraphPropertyState({ type: "imported", token: state.token, authoredPPr });
    case "editor-created":
      return trustParagraphPropertyState({ type: "editor-created", authoredPPr });
    case "transient-template":
      return trustParagraphPropertyState({
        type: "transient-template",
        handle: state.handle,
        authoredPPr,
      });
    default: {
      const exhaustive: never = state;
      return exhaustive;
    }
  }
};

export const applyParagraphPropertyMutations = (
  authoredPPr: AuthoredParagraphProperties,
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
        left: createEditorParagraphPropertyState(state.authoredPPr),
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

export const EMPTY_EDITOR_PARAGRAPH_PROPERTY_STATE = Object.freeze({
  type: "editor-created",
  authoredPPr: Object.freeze({}),
}) satisfies SerializedPersistableParagraphPropertyState;
