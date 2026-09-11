import { panic } from "better-result";
import type { Fragment, Mark, Node as PMNode, NodeType } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { setProseParagraphMarkupWithPropertySource } from "../docx/paragraphPropertySource";
import type { ParagraphFormatting } from "../types/document";
import { mergeParagraphFormatting } from "../utils/paragraphFormattingMerge";
import { directionFromBidi, directionToBidi } from "./paragraphDirection";
import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  PARAGRAPH_FORMATTING_PROPERTY_ATTRS,
  PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST,
  PARAGRAPH_MARK_ATTR_KEYS,
  PPR_CHANGE_SCOPED_ATTR_KEYS,
  type AuthoredParagraphProperties,
} from "./paragraphPropertyProjection";
import {
  authoredParagraphPropertiesFromFormatting,
  createEditorParagraphPropertyState,
  expectParagraphPropertyState,
  readAuthoredParagraphProperties,
  serializeParagraphPropertyState,
  serializePersistableParagraphPropertyState,
  transitionParagraphPropertyState,
  type ParagraphPropertyState,
  type ParagraphPropertyStateTransition,
} from "./paragraphPropertyState";
import { paragraphSpacingAttrPatch, paragraphSpacingFromFormatting } from "./paragraphSpacing";
import { autospacingMatchesBase } from "./autospacingBase";
import type { ParagraphAttrs } from "./schema/nodes";

const deepFreezeProjectionValue = (value: unknown): unknown => {
  if (Array.isArray(value)) {
    for (const entry of value) {
      deepFreezeProjectionValue(entry);
    }
    return Object.freeze(value);
  }
  if (typeof value === "object" && value !== null) {
    for (const entry of Object.values(value)) {
      deepFreezeProjectionValue(entry);
    }
    return Object.freeze(value);
  }
  return value;
};

/** Nominal capsule proving governed attrs and authored state were projected together. */
class ParagraphPropertyProjectionCapsule {
  readonly #attrs: ParagraphAttrs;

  private constructor(attrs: ParagraphAttrs) {
    this.#attrs = deepFreezeProjectionValue(structuredClone(attrs)) as ParagraphAttrs;
    Object.freeze(this);
  }

  static fromProjectedAttrs(attrs: ParagraphAttrs): ParagraphPropertyProjectionCapsule {
    return new ParagraphPropertyProjectionCapsule(attrs);
  }

  static attrs(projection: ParagraphPropertyProjectionCapsule): ParagraphAttrs {
    return projection.#attrs;
  }
}

export type ParagraphPropertyProjection = ParagraphPropertyProjectionCapsule;

export type ParagraphPropertyProjectionContext = {
  inheritedPPr: AuthoredParagraphProperties;
  numberingLevelIndent: ParagraphFormatting["numberingLevelIndent"] | null;
  numPrFromStyle: ParagraphFormatting["numPr"] | null;
};

export type ParagraphPropertyContextTransition =
  | { type: "preserve" }
  | ({ type: "replace" } & ParagraphPropertyProjectionContext);

type TransitionParagraphPropertiesOptions = {
  attrs: ParagraphAttrs;
  state: ParagraphPropertyStateTransition;
  context: ParagraphPropertyContextTransition;
};

type CreateParagraphPropertiesOptions = {
  attrs?: Readonly<Record<string, unknown>>;
  state: ParagraphPropertyState;
  context: ParagraphPropertyProjectionContext;
  paragraphMarkFormatting?: ParagraphAttrs["_paragraphMarkFormatting"];
};

type CreateEditorParagraphPropertiesOptions = Omit<
  CreateParagraphPropertiesOptions,
  "state"
> & {
  authoredPPr: AuthoredParagraphProperties;
};

const PROJECTION_CONTEXT_ATTR_KEYS = [
  "_paragraphPropertyState",
  "_paragraphPropertyInheritance",
  "_paragraphMarkFormatting",
  "_numberingLevelIndent",
  "numPrFromStyle",
  "alignmentFromStyle",
  "_sourceIndentation",
  "_autospacingBase",
  "spacingFromDocDefaults",
  "spacingFromImplicitDefaultStyle",
] as const satisfies readonly (keyof ParagraphAttrs)[];

export const GOVERNED_PARAGRAPH_ATTR_KEYS: ReadonlySet<keyof ParagraphAttrs> = new Set([
  ...PPR_CHANGE_SCOPED_ATTR_KEYS,
  ...PARAGRAPH_MARK_ATTR_KEYS,
  ...PROJECTION_CONTEXT_ATTR_KEYS,
]);

/** Select exact authored pPr from editor attrs through the total descriptor. */
const authoredParagraphPropertiesFromExternalDomAttrs = (
  attrs: Readonly<Partial<ParagraphAttrs>>,
): AuthoredParagraphProperties => {
  const authored: ParagraphFormatting = {};
  for (const rawKey of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
    const descriptor = PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[rawKey];
    if (descriptor.owner !== "pPr-base") {
      continue;
    }
    switch (descriptor.projection) {
      case "alignment":
        if (attrs.alignment !== undefined && attrs.alignment !== null) {
          authored.alignment = attrs.alignment;
        }
        break;
      case "direction": {
        const bidi = directionToBidi(attrs.direction);
        if (bidi !== undefined) {
          authored.bidi = bidi;
        }
        break;
      }
      case "direct":
      case "boolean":
      case "opaque": {
        const attr = PARAGRAPH_FORMATTING_PROPERTY_ATTRS[rawKey].at(0);
        const value = attr === undefined ? undefined : Reflect.get(attrs, attr);
        if (value !== undefined && value !== null) {
          Reflect.set(authored, rawKey, structuredClone(value));
        }
        break;
      }
      case "indentation": {
        const attr = PARAGRAPH_FORMATTING_PROPERTY_ATTRS[rawKey].at(0);
        const value = attr === undefined ? undefined : Reflect.get(attrs, attr);
        if (value !== undefined && value !== null) {
          Reflect.set(authored, rawKey, value);
        }
        break;
      }
      case "spacing":
        if (rawKey === "beforeAutospacing") {
          if (autospacingMatchesBase(attrs._autospacingBase, "before", attrs.spaceBefore)) {
            authored.beforeAutospacing = true;
          }
        } else if (rawKey === "afterAutospacing") {
          if (autospacingMatchesBase(attrs._autospacingBase, "after", attrs.spaceAfter)) {
            authored.afterAutospacing = true;
          }
        } else if (rawKey === "spacingExplicit") {
          if (attrs.spacingExplicit !== undefined && attrs.spacingExplicit !== null) {
            authored.spacingExplicit = structuredClone(attrs.spacingExplicit);
          }
        } else {
          const attr = PARAGRAPH_FORMATTING_PROPERTY_ATTRS[rawKey].at(0);
          const value = attr === undefined ? undefined : Reflect.get(attrs, attr);
          if (value !== undefined && value !== null) {
            Reflect.set(authored, rawKey, value);
          }
        }
        break;
      case "numbering-provenance":
      case "preserve-live":
        break;
      default: {
        const exhaustive: never = descriptor.projection;
        return exhaustive;
      }
    }
  }
  return authoredParagraphPropertiesFromFormatting(authored);
};

const nonGovernedAttrs = (
  attrs: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> => {
  const result: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (!GOVERNED_PARAGRAPH_ATTR_KEYS.has(key as keyof ParagraphAttrs)) {
      result[key] = value;
    }
  }
  return result;
};

type CreateEditorParagraphPropertiesFromAttrsOptions = {
  effectiveAttrs: Readonly<Partial<ParagraphAttrs>>;
  authoredAttrs?: Readonly<Partial<ParagraphAttrs>>;
  inheritedPPr?: AuthoredParagraphProperties;
};

const createEditorParagraphPropertiesFromExternalDomAttrs = ({
  effectiveAttrs,
  authoredAttrs = effectiveAttrs,
  inheritedPPr = {},
}: CreateEditorParagraphPropertiesFromAttrsOptions): ParagraphPropertyProjection =>
  createEditorParagraphProperties({
    attrs: nonGovernedAttrs(effectiveAttrs),
    authoredPPr: authoredParagraphPropertiesFromExternalDomAttrs(authoredAttrs),
    context: { inheritedPPr, numberingLevelIndent: null, numPrFromStyle: null },
    paragraphMarkFormatting: effectiveAttrs._paragraphMarkFormatting ?? {},
  });

const cloneFormatting = (formatting: Readonly<AuthoredParagraphProperties>): ParagraphFormatting =>
  structuredClone(formatting);

const sourceIndentation = (
  direct: Readonly<AuthoredParagraphProperties>,
  inherited: ParagraphFormatting | undefined,
): ParagraphAttrs["_sourceIndentation"] | null => {
  const result: NonNullable<ParagraphAttrs["_sourceIndentation"]> = {};
  if (direct.indentLeft === undefined && inherited?.indentLeft !== undefined) {
    result.indentLeft = inherited.indentLeft;
  }
  if (direct.indentRight === undefined && inherited?.indentRight !== undefined) {
    result.indentRight = inherited.indentRight;
  }
  const ownsFirstLine = direct.indentFirstLine !== undefined || direct.hangingIndent !== undefined;
  if (!ownsFirstLine && inherited?.indentFirstLine !== undefined) {
    result.indentFirstLine = inherited.indentFirstLine;
    result.hangingIndent = inherited.hangingIndent === true;
  }
  return Object.keys(result).length > 0 ? result : null;
};

const numberingIndentFormatting = (
  provenance: ParagraphFormatting["numberingLevelIndent"] | null,
): ParagraphFormatting | undefined => {
  if (provenance?.type !== "owned") {
    return undefined;
  }
  return {
    ...(provenance.owned.indentLeft !== undefined
      ? { indentLeft: provenance.owned.indentLeft }
      : {}),
    ...(provenance.owned.indentRight !== undefined
      ? { indentRight: provenance.owned.indentRight }
      : {}),
    ...(provenance.owned.indentFirstLine !== undefined
      ? { indentFirstLine: provenance.owned.indentFirstLine }
      : {}),
    ...(provenance.owned.hangingIndent !== undefined
      ? { hangingIndent: provenance.owned.hangingIndent }
      : {}),
  };
};

const projectionContext = (
  attrs: ParagraphAttrs,
  transition: ParagraphPropertyContextTransition,
): ParagraphPropertyProjectionContext => {
  switch (transition.type) {
    case "preserve":
      return {
        inheritedPPr: attrs._paragraphPropertyInheritance,
        numberingLevelIndent: attrs._numberingLevelIndent ?? null,
        numPrFromStyle: attrs.numPrFromStyle ?? null,
      };
    case "replace":
      return {
        inheritedPPr: authoredParagraphPropertiesFromFormatting(transition.inheritedPPr),
        numberingLevelIndent: transition.numberingLevelIndent,
        numPrFromStyle: transition.numPrFromStyle,
      };
    default: {
      const exhaustive: never = transition;
      return exhaustive;
    }
  }
};

const projectAttrs = (
  attrs: ParagraphAttrs,
  state: ParagraphPropertyState,
  context: ParagraphPropertyProjectionContext,
): ParagraphAttrs => {
  const authoredPPr = cloneFormatting(state.authoredPPr);
  const inheritedPPr = cloneFormatting(context.inheritedPPr);
  const inheritedOverNumbering = mergeParagraphFormatting(
    numberingIndentFormatting(context.numberingLevelIndent),
    inheritedPPr,
  );
  const effective = mergeParagraphFormatting(inheritedOverNumbering, authoredPPr) ?? {};
  const next: ParagraphAttrs = {
    ...attrs,
    _paragraphPropertyState: serializeParagraphPropertyState(state),
    _paragraphPropertyInheritance: authoredParagraphPropertiesFromFormatting(inheritedPPr),
  };
  Reflect.set(next, "_numberingLevelIndent", context.numberingLevelIndent ?? null);
  Reflect.set(next, "numPrFromStyle", context.numPrFromStyle ?? null);

  Object.assign(
    next,
    paragraphSpacingAttrPatch({
      direct: paragraphSpacingFromFormatting(authoredPPr),
      inherited: paragraphSpacingFromFormatting(inheritedPPr),
    }),
  );
  if (authoredPPr.spacingExplicit !== undefined) {
    next.spacingExplicit = structuredClone(authoredPPr.spacingExplicit);
  }

  for (const rawKey of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
    const descriptor = PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[rawKey];
    switch (descriptor.projection) {
      case "alignment":
        Reflect.set(next, "alignment", effective.alignment ?? null);
        Reflect.set(next, "alignmentFromStyle", inheritedPPr.alignment ?? null);
        break;
      case "direction":
        next.direction = directionFromBidi(effective.bidi);
        break;
      case "direct":
      case "boolean":
      case "opaque":
        for (const attr of PARAGRAPH_FORMATTING_PROPERTY_ATTRS[rawKey]) {
          Reflect.set(next, attr, Reflect.get(effective, rawKey) ?? null);
        }
        break;
      case "indentation":
      case "spacing":
      case "numbering-provenance":
      case "preserve-live":
        break;
      default: {
        const exhaustive: never = descriptor.projection;
        return exhaustive;
      }
    }
  }

  Reflect.set(next, "indentLeft", effective.indentLeft ?? null);
  Reflect.set(next, "indentRight", effective.indentRight ?? null);
  Reflect.set(next, "indentFirstLine", effective.indentFirstLine ?? null);
  Reflect.set(
    next,
    "hangingIndent",
    effective.indentFirstLine === undefined ? null : (effective.hangingIndent ?? false),
  );
  Reflect.set(next, "_sourceIndentation", sourceIndentation(authoredPPr, inheritedOverNumbering));
  return next;
};

const proveParagraphPropertyProjection = (
  attrs: ParagraphAttrs,
): ParagraphPropertyProjection => ParagraphPropertyProjectionCapsule.fromProjectedAttrs(attrs);

export const transitionParagraphProperties = ({
  attrs,
  state: transition,
  context: contextTransition,
}: TransitionParagraphPropertiesOptions): ParagraphPropertyProjection => {
  const current = expectParagraphPropertyState(attrs._paragraphPropertyState);
  const state = transitionParagraphPropertyState(current, transition);
  const context = projectionContext(attrs, contextTransition);
  return proveParagraphPropertyProjection(projectAttrs(attrs, state, context));
};

export const createParagraphProperties = ({
  attrs = {},
  state,
  context,
  paragraphMarkFormatting = {},
}: CreateParagraphPropertiesOptions): ParagraphPropertyProjection => {
  for (const key of Object.keys(attrs)) {
    if (GOVERNED_PARAGRAPH_ATTR_KEYS.has(key as keyof ParagraphAttrs)) {
      panic(`Editor paragraph creation supplied governed attr directly: ${key}`);
    }
  }
  const seed = {
    ...attrs,
    _paragraphPropertyState: serializeParagraphPropertyState(state),
    _paragraphPropertyInheritance: context.inheritedPPr,
    _paragraphMarkFormatting: paragraphMarkFormatting,
  } as ParagraphAttrs;
  return proveParagraphPropertyProjection(projectAttrs(seed, state, context));
};

export const createEditorParagraphProperties = ({
  attrs,
  authoredPPr,
  context,
  paragraphMarkFormatting,
}: CreateEditorParagraphPropertiesOptions): ParagraphPropertyProjection =>
  createParagraphProperties({
    attrs,
    state: createEditorParagraphPropertyState(authoredPPr),
    context,
    paragraphMarkFormatting,
  });

const valuesEqual = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) {
    return true;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    return left.every((entry, index) => valuesEqual(entry, right[index]));
  }
  if (
    typeof left !== "object" ||
    left === null ||
    typeof right !== "object" ||
    right === null
  ) {
    return false;
  }
  const leftEntries = Object.entries(left);
  const rightRecord = right as Record<string, unknown>;
  if (leftEntries.length !== Object.keys(rightRecord).length) {
    return false;
  }
  return leftEntries.every(
    ([key, value]) => Object.hasOwn(rightRecord, key) && valuesEqual(value, rightRecord[key]),
  );
};

export type ParagraphPropertyInvariantBoundary =
  | { type: "internal" }
  | { type: "persistence" };

/** Prove state, context, and every descriptor-governed effective attr are a fixed point. */
export const assertParagraphPropertyInvariant = (
  attrs: ParagraphAttrs,
  boundary: ParagraphPropertyInvariantBoundary,
): ParagraphPropertyState => {
  const state = expectParagraphPropertyState(attrs._paragraphPropertyState);
  if (boundary.type === "persistence") {
    serializePersistableParagraphPropertyState(state);
  }
  const inheritance = readAuthoredParagraphProperties(attrs._paragraphPropertyInheritance);
  if (inheritance.status !== "valid") {
    panic("Paragraph-property inheritance must be a valid authored-property projection");
  }
  const expected = projectAttrs(attrs, state, {
    inheritedPPr: inheritance.value,
    numberingLevelIndent: attrs._numberingLevelIndent ?? null,
    numPrFromStyle: attrs.numPrFromStyle ?? null,
  });
  for (const key of GOVERNED_PARAGRAPH_ATTR_KEYS) {
    if (!valuesEqual(attrs[key], expected[key])) {
      panic(`Paragraph-property projection invariant failed for attr: ${String(key)}`);
    }
  }
  return state;
};

/** Create proof for a mutation that leaves every governed property unchanged. */
export const preserveParagraphProperties = (
  attrs: ParagraphAttrs,
  patch: Readonly<Record<string, unknown>>,
): ParagraphPropertyProjection => {
  const next = { ...attrs, ...patch } as ParagraphAttrs;
  for (const key of GOVERNED_PARAGRAPH_ATTR_KEYS) {
    if (!valuesEqual(attrs[key], next[key])) {
      panic(`Non-governed paragraph mutation changed governed attr: ${String(key)}`);
    }
  }
  expectParagraphPropertyState(next._paragraphPropertyState);
  return proveParagraphPropertyProjection(next);
};

type ApplyParagraphPropertyProjectionOptions = {
  transaction: Transaction;
  pos: number;
  projection: ParagraphPropertyProjection;
  source: { type: "preserve" } | { type: "transfer-allocated-id" };
};

/** The only paragraph-markup mutation sink accepted by the ownership guard. */
export const applyParagraphPropertyProjection = ({
  transaction,
  pos,
  projection,
  source,
}: ApplyParagraphPropertyProjectionOptions): void => {
  const attrs = ParagraphPropertyProjectionCapsule.attrs(projection);
  assertParagraphPropertyInvariant(attrs, { type: "internal" });
  setProseParagraphMarkupWithPropertySource({
    transaction,
    pos,
    attrs,
    ownership: source.type,
  });
};

type CreateParagraphNodeFromProjectionOptions = {
  type: NodeType;
  projection: ParagraphPropertyProjection;
  content?: Fragment | PMNode | readonly PMNode[] | null;
  marks?: readonly Mark[];
};

/** The only paragraph-node creation sink accepted by the ownership guard. */
export const createParagraphNodeFromProjection = ({
  type,
  projection,
  content,
  marks,
}: CreateParagraphNodeFromProjectionOptions): PMNode => {
  if (type.name !== "paragraph") {
    panic("Paragraph-property projection can only create a paragraph node");
  }
  const attrs = ParagraphPropertyProjectionCapsule.attrs(projection);
  assertParagraphPropertyInvariant(attrs, { type: "internal" });
  return type.create(attrs, content, marks);
};

type ApplyNonParagraphMarkupOptions = {
  transaction: Transaction;
  pos: number;
  attrs?: PMNode["attrs"];
  type?: NodeType;
  marks?: readonly Mark[];
};

/** Canonical markup sink for nodes that are runtime-proven not to be paragraphs. */
export const applyNonParagraphMarkup = ({
  transaction,
  pos,
  attrs,
  type,
  marks,
}: ApplyNonParagraphMarkupOptions): void => {
  const current = transaction.doc.nodeAt(pos);
  if (!current) {
    panic("Cannot mutate markup for a missing node");
  }
  if (current.type.name === "paragraph" || type?.name === "paragraph") {
    panic("Non-paragraph markup sink cannot mutate a paragraph");
  }
  transaction.setNodeMarkup(pos, type, attrs, marks);
};

/** Explicit external-DOM import boundary used by the paragraph NodeSpec parser. */
export const paragraphAttrsFromExternalDomImport = ({
  effectiveAttrs,
  authoredAttrs = effectiveAttrs,
  inheritedPPr = {},
}: CreateEditorParagraphPropertiesFromAttrsOptions): ParagraphAttrs => {
  const projection = createEditorParagraphPropertiesFromExternalDomAttrs({
    effectiveAttrs,
    authoredAttrs,
    inheritedPPr,
  });
  return ParagraphPropertyProjectionCapsule.attrs(projection);
};
