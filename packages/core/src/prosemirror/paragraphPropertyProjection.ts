import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST,
  PARAGRAPH_PROPERTY_OWNER,
  PARAGRAPH_PROPERTY_PROJECTION,
  PARAGRAPH_STYLE_TRANSITION,
  type ParagraphPropertyDescriptor,
} from "../docx/paragraphPropertyDescriptor";
import type { ParagraphAttrs } from "./schema/nodes";

export * from "../docx/paragraphPropertyDescriptor";

type ParagraphFormattingPropertyKey = keyof typeof PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR;

/**
 * Total ProseMirror projection of the representation-neutral formatting model.
 *
 * This companion belongs to the PM adapter: the model-level descriptor must not
 * depend on schema attr names, while `satisfies` makes both a new model property
 * and an attr rename fail compilation here.
 */
export const PARAGRAPH_FORMATTING_PROPERTY_ATTRS = {
  alignment: ["alignment"],
  bidi: ["direction"],
  kinsoku: ["kinsoku"],
  overflowPunctuation: ["overflowPunctuation"],
  spaceBefore: ["spaceBefore"],
  spaceAfter: ["spaceAfter"],
  lineSpacing: ["lineSpacing"],
  lineSpacingRule: ["lineSpacingRule", "lineSpacingExplicit"],
  snapToGrid: ["snapToGrid"],
  beforeAutospacing: ["_autospacingBase"],
  afterAutospacing: ["_autospacingBase"],
  spacingExplicit: ["spacingExplicit"],
  indentLeft: ["indentLeft"],
  indentRight: ["indentRight"],
  indentFirstLine: ["indentFirstLine"],
  hangingIndent: ["hangingIndent"],
  numberingLevelIndent: [],
  borders: ["borders"],
  shading: ["shading"],
  tabs: ["tabs"],
  keepNext: ["keepNext"],
  keepLines: ["keepLines"],
  widowControl: ["widowControl"],
  pageBreakBefore: ["pageBreakBefore"],
  contextualSpacing: ["contextualSpacing"],
  numPr: ["numPr"],
  numPrFromStyle: ["numPrFromStyle"],
  outlineLevel: ["outlineLevel"],
  styleId: ["styleId"],
  frame: ["frame"],
  suppressLineNumbers: ["suppressLineNumbers"],
  suppressAutoHyphens: ["suppressAutoHyphens"],
  runProperties: ["defaultTextFormatting"],
  runInWithNext: ["runInWithNext"],
} as const satisfies Record<
  ParagraphFormattingPropertyKey,
  readonly (keyof ParagraphAttrs)[]
>;

type ParagraphPropertyDescriptorAttr =
  (typeof PARAGRAPH_FORMATTING_PROPERTY_ATTRS)[ParagraphFormattingPropertyKey][number];

const propertyAttrs = (
  predicate: (descriptor: ParagraphPropertyDescriptor) => boolean,
): readonly ParagraphPropertyDescriptorAttr[] => {
  const attrs = new Set<ParagraphPropertyDescriptorAttr>();
  for (const rawKey of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
    const descriptor = PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[rawKey];
    if (!predicate(descriptor)) {
      continue;
    }
    for (const attr of PARAGRAPH_FORMATTING_PROPERTY_ATTRS[rawKey]) {
      attrs.add(attr);
    }
  }
  return Object.freeze([...attrs]);
};

/** Paragraph attrs governed by a `w:pPrChange` CT_PPrBase payload. */
export const PPR_CHANGE_SCOPED_ATTR_KEYS = propertyAttrs(
  ({ owner }) => owner === PARAGRAPH_PROPERTY_OWNER.pPrBase,
);

/** The direct boolean attr class serialized as OOXML on/off properties. */
export const PPR_CHANGE_BOOLEAN_ATTR_KEYS = propertyAttrs(
  ({ owner, projection }) =>
    owner === PARAGRAPH_PROPERTY_OWNER.pPrBase &&
    projection === PARAGRAPH_PROPERTY_PROJECTION.boolean,
);

/** Paragraph-mark attrs governed beside, but not inside, CT_PPrBase state. */
export const PARAGRAPH_MARK_ATTR_KEYS = propertyAttrs(
  ({ owner }) => owner === PARAGRAPH_PROPERTY_OWNER.paragraphMark,
);

export const PPR_SPACING_ATTR_KEY_LIST = Object.freeze(
  propertyAttrs(({ projection }) => projection === PARAGRAPH_PROPERTY_PROJECTION.spacing),
);
const PPR_SPACING_ATTR_KEY_SET = new Set<keyof ParagraphAttrs>(PPR_SPACING_ATTR_KEY_LIST);
export const isPprSpacingAttr = (key: keyof ParagraphAttrs): boolean =>
  PPR_SPACING_ATTR_KEY_SET.has(key);

export const PPR_INDENT_ATTR_KEY_LIST = Object.freeze(
  propertyAttrs(({ projection }) => projection === PARAGRAPH_PROPERTY_PROJECTION.indentation),
);
const PPR_INDENT_ATTR_KEY_SET = new Set<keyof ParagraphAttrs>(PPR_INDENT_ATTR_KEY_LIST);
export const isPprIndentAttr = (key: keyof ParagraphAttrs): boolean =>
  PPR_INDENT_ATTR_KEY_SET.has(key);

export const PPR_STYLE_REPLACED_ATTR_KEYS = propertyAttrs(
  ({ styleTransition }) => styleTransition === PARAGRAPH_STYLE_TRANSITION.replace,
);
