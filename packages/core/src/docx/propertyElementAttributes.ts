/**
 * Which attributes of a property element the model has a field for.
 *
 * A property element is an attribute bag: `w:ind`, `w:spacing`, `w:framePr`,
 * `w:tab`, a `w:pBdr` side, `w:shd`. The child dispatcher decides them whole —
 * a handler either reads the element or hands back its bytes — and that is
 * exactly as far as it goes. `<w:ind w:leftChars="100"/>` is kept entire
 * because the reader took nothing from it, and `<w:ind w:left="720"
 * w:leftChars="100"/>` is modelled and loses the character unit, because the
 * reader took something. The census measures one attribute at a time, so the
 * second case reported as surviving.
 *
 * The fix is the attribute remainder that `w:p`'s `w:rsid*` attributes already
 * ride, applied one level down: every attribute with no model field is kept on
 * the record, decided on the resolved local name, written back by
 * `serializePreservedAttributes`. The remainder predicate has to be derived
 * from the model rather than hand-listed, or it is a mirror of the reader and
 * drifts the first time a field is added — so each table below is
 * `as const satisfies ModelledAttributes<…>` over the record's own fields, and
 * a field that gains no entry does not compile.
 *
 * The tables live together, in a module with no runtime dependencies, because
 * three callers consult them: the reader (to compute the remainder), and the
 * container-survival census (to state a subject attribute beside a modelled
 * sibling, which is what catches a whole-or-nothing reader at all). The writer
 * needs no copy: it hands the fragments it is about to emit to
 * `serializePreservedAttributes`, which drops a remainder entry that would
 * spell one of them again.
 *
 * A field may name more than one attribute. `ParagraphFormatting.indentLeft`
 * is filled from `w:left` or from the Strict `w:start`, and both have to be
 * out of the remainder or a save writes the same indent twice under two
 * spellings.
 */

import type {
  BorderSpec,
  ParagraphFormatting,
  ShadingProperties,
  TabStop,
} from "../types/document";

/**
 * The attributes one record's fields are read from, keyed by field.
 *
 * `Record`, never `Partial<Record>`: the key set is the record's own fields,
 * so a field that has no attribute to name has to say so with an empty list
 * rather than by being absent.
 */
export type ModelledAttributes<Field extends string> = Readonly<
  Record<Field, string | readonly string[]>
>;

/** A record's attribute-bearing fields: all of them but the remainder itself. */
type AttributeFields<Shape> = Exclude<keyof Required<Shape>, "preservedAttributes"> & string;

type IndentationField = Extract<
  keyof ParagraphFormatting,
  "indentLeft" | "indentRight" | "indentFirstLine" | "hangingIndent"
>;

/**
 * `w:ind`. `hangingIndent` is the sign of `indentFirstLine` rather than a
 * value of its own, so both name the attribute pair that carries it.
 */
export const INDENTATION_ATTRIBUTES = {
  indentLeft: ["left", "start"],
  indentRight: ["right", "end"],
  indentFirstLine: ["firstLine", "hanging"],
  hangingIndent: "hanging",
} as const satisfies ModelledAttributes<IndentationField>;

type SpacingField = Extract<
  keyof ParagraphFormatting,
  | "spaceBefore"
  | "spaceAfter"
  | "lineSpacing"
  | "lineSpacingRule"
  | "beforeAutospacing"
  | "afterAutospacing"
  | "spacingExplicit"
>;

/**
 * `w:spacing`. `spacingExplicit` records which of the two sides the paragraph
 * stated itself; it is derived from them and reads no attribute of its own.
 */
export const SPACING_ATTRIBUTES = {
  spaceBefore: "before",
  spaceAfter: "after",
  lineSpacing: "line",
  lineSpacingRule: "lineRule",
  beforeAutospacing: "beforeAutospacing",
  afterAutospacing: "afterAutospacing",
  spacingExplicit: [],
} as const satisfies ModelledAttributes<SpacingField>;

export const FRAME_ATTRIBUTES = {
  dropCap: "dropCap",
  lines: "lines",
  width: "w",
  height: "h",
  hSpace: "hSpace",
  vSpace: "vSpace",
  hAnchor: "hAnchor",
  vAnchor: "vAnchor",
  x: "x",
  y: "y",
  xAlign: "xAlign",
  yAlign: "yAlign",
  wrap: "wrap",
} as const satisfies ModelledAttributes<AttributeFields<NonNullable<ParagraphFormatting["frame"]>>>;

export const TAB_STOP_ATTRIBUTES = {
  position: "pos",
  alignment: "val",
  leader: "leader",
} as const satisfies ModelledAttributes<AttributeFields<TabStop>>;

/**
 * `CT_Border`. The five art relationship ids are in the relationships
 * namespace; the remainder decides on the local name, so naming them here
 * keeps a `r:topLeft` out of it.
 */
export const BORDER_ATTRIBUTES = {
  style: "val",
  color: ["color", "themeColor", "themeTint", "themeShade"],
  size: "sz",
  space: "space",
  shadow: "shadow",
  frame: "frame",
  artRelationshipId: "id",
  topLeftArtRelationshipId: "topLeft",
  topRightArtRelationshipId: "topRight",
  bottomLeftArtRelationshipId: "bottomLeft",
  bottomRightArtRelationshipId: "bottomRight",
} as const satisfies ModelledAttributes<AttributeFields<BorderSpec>>;

/** `CT_Shd`: direct, theme, modifier, and pattern attributes all have model fields. */
export const SHADING_ATTRIBUTES = {
  color: ["color", "themeColor", "themeTint", "themeShade"],
  fill: ["fill", "themeFill", "themeFillTint", "themeFillShade"],
  pattern: "val",
} as const satisfies ModelledAttributes<AttributeFields<ShadingProperties>>;

/**
 * The same tables by the complex type whose attributes they answer for, in
 * the order the model declares the fields.
 *
 * The census reads this: to catch a whole-or-nothing reader it has to state
 * the subject attribute *beside* one the element models, and which one that
 * is has to come from the model rather than from a list in the census. The
 * first entry is the one it picks.
 */
export const PROPERTY_ELEMENT_ATTRIBUTES = {
  CT_Ind: INDENTATION_ATTRIBUTES,
  CT_Spacing: SPACING_ATTRIBUTES,
  CT_FramePr: FRAME_ATTRIBUTES,
  CT_TabStop: TAB_STOP_ATTRIBUTES,
  CT_Border: BORDER_ATTRIBUTES,
  CT_Shd: SHADING_ATTRIBUTES,
} as const satisfies Readonly<Record<string, ModelledAttributes<string>>>;

/** Every attribute local name a table names, flattened, in field order. */
export const modelledAttributeNames = (modelled: ModelledAttributes<string>): readonly string[] => {
  const names: string[] = [];
  for (const declared of Object.values(modelled)) {
    if (typeof declared === "string") {
      names.push(declared);
      continue;
    }
    names.push(...declared);
  }
  return names;
};
