/**
 * Shared mark utility functions
 *
 * setMark, removeMark, isMarkActive, getMarkAttr, marksToTextFormatting,
 * textFormattingToMarks, clearFormatting
 */

import type { MarkType, Mark, Schema } from "prosemirror-model";
import { toggleMark } from "prosemirror-commands";
import type { Command, EditorState, Transaction } from "prosemirror-state";

import type { TextFormatting, UnderlineStyle, ThemeColorSlot } from "../../../types/document";
import { FONT_THEME_VALUES } from "../../../types/documentEnumValues";
import { mergeFontFamily } from "../../../utils/fontFamilyMerge";
import { expectFontFamilyMarkAttrs, expectRunFormattingOverrideMarkAttrs } from "../../attrs";
import { selectRunFormattingCarrierRepresentations } from "../../runFormattingInlineCarriers";
import { hasRunFormattingOverrideAttrs } from "../../runFormattingProvenance";
import { normalizeHorizontalScalePercent } from "../../../utils/horizontalScale";
import { shadingToRunShadingAttrs } from "../../conversion/runShadingMark";
import {
  COMPLEX_SCRIPT_RUN_PROPERTY_KEYS,
  type ComplexScriptRunPropertyKey,
} from "../../schema/marks";
import type { FontFamilyAttrs, RunFormattingOverrideAttrs } from "../../schema/marks";
import {
  applyRunFormattingOverrideMark,
  buildRunFormattingOverrideAttrs,
} from "./RunFormattingOverrideExtension";

type MarkAttrs = Record<string, unknown>;
type FontFamilyFormatting = NonNullable<TextFormatting["fontFamily"]>;
type FontTheme = NonNullable<FontFamilyFormatting["asciiTheme"]>;

const isFontTheme = (value: string | undefined): value is FontTheme =>
  value !== undefined && FONT_THEME_VALUES.some((theme) => theme === value);

const fontFamilyAttrsToFormatting = ({
  ascii,
  hAnsi,
  eastAsia,
  cs,
  hint,
  asciiTheme,
  hAnsiTheme,
  eastAsiaTheme,
  csTheme,
}: FontFamilyAttrs): FontFamilyFormatting => ({
  ...(ascii !== undefined ? { ascii } : {}),
  ...(hAnsi !== undefined ? { hAnsi } : {}),
  ...(eastAsia !== undefined ? { eastAsia } : {}),
  ...(cs !== undefined ? { cs } : {}),
  ...(hint !== undefined ? { hint } : {}),
  ...(isFontTheme(asciiTheme) ? { asciiTheme } : {}),
  ...(hAnsiTheme !== undefined ? { hAnsiTheme } : {}),
  ...(eastAsiaTheme !== undefined ? { eastAsiaTheme } : {}),
  ...(csTheme !== undefined ? { csTheme } : {}),
});

// ============================================================================
// PARAGRAPH DEFAULT FORMATTING HELPERS
// ============================================================================

function marksToTextFormatting(marks: readonly Mark[]): TextFormatting {
  const formatting: TextFormatting = {};

  for (const mark of marks) {
    switch (mark.type.name) {
      case "bold":
        formatting.bold = true;
        break;
      case "italic":
        formatting.italic = true;
        break;
      case "underline": {
        // SAFETY: underline mark always has style attr per schema; value is a valid UnderlineStyle
        const underlineStyle = (
          typeof mark.attrs["style"] === "string" ? mark.attrs["style"] : "single"
        ) as UnderlineStyle;
        formatting.underline = {
          style: underlineStyle,
          ...(mark.attrs["color"] !== null && mark.attrs["color"] !== undefined
            ? { color: mark.attrs["color"] }
            : {}),
        };
        break;
      }
      case "strike":
        formatting.strike = true;
        break;
      case "textColor": {
        // SAFETY: textColor mark attrs always match ColorValue shape — extracted individually;
        // themeColor is always a valid ThemeColorSlot string per schema
        const colorRgb =
          mark.attrs["rgb"] !== null && mark.attrs["rgb"] !== undefined
            ? String(mark.attrs["rgb"])
            : undefined;
        const colorTheme =
          mark.attrs["themeColor"] !== null && mark.attrs["themeColor"] !== undefined
            ? (String(mark.attrs["themeColor"]) as ThemeColorSlot)
            : undefined;
        const colorTint =
          mark.attrs["themeTint"] !== null && mark.attrs["themeTint"] !== undefined
            ? String(mark.attrs["themeTint"])
            : undefined;
        const colorShade =
          mark.attrs["themeShade"] !== null && mark.attrs["themeShade"] !== undefined
            ? String(mark.attrs["themeShade"])
            : undefined;
        formatting.color = {
          ...(colorRgb !== undefined ? { rgb: colorRgb } : {}),
          ...(colorTheme !== undefined ? { themeColor: colorTheme } : {}),
          ...(colorTint !== undefined ? { themeTint: colorTint } : {}),
          ...(colorShade !== undefined ? { themeShade: colorShade } : {}),
        };
        break;
      }
      case "highlight":
        // SAFETY: highlight mark always has color attr per schema; value is a valid highlight union member
        formatting.highlight = String(mark.attrs["color"]) as NonNullable<
          TextFormatting["highlight"]
        >;
        break;
      case "fontSize":
        // SAFETY: fontSize mark always has size attr per schema
        formatting.fontSize = Number(mark.attrs["size"]);
        break;
      case "fontFamily": {
        formatting.fontFamily = fontFamilyAttrsToFormatting(expectFontFamilyMarkAttrs(mark));
        break;
      }
      case "language": {
        const val = mark.attrs["val"];
        const eastAsia = mark.attrs["eastAsia"];
        const bidi = mark.attrs["bidi"];
        formatting.language = {
          ...(typeof val === "string" ? { val } : {}),
          ...(typeof eastAsia === "string" ? { eastAsia } : {}),
          ...(typeof bidi === "string" ? { bidi } : {}),
        };
        break;
      }
      case "superscript":
        formatting.vertAlign = "superscript";
        break;
      case "subscript":
        formatting.vertAlign = "subscript";
        break;
      case "rtl":
        // eigenpal/docx-editor#806 — keep per-run RTL direction through the
        // live-edit/clipboard/keymap mark paths (the `rtl=false` negative
        // override rides `runFormattingOverride`, so only the `true` case
        // needs an explicit branch here).
        formatting.rtl = true;
        break;
      case "runFormattingOverride":
        applyRunFormattingOverrideMark(formatting, mark);
        break;
      default:
        break;
    }
  }

  return formatting;
}

function saveStoredMarksToParagraph(
  state: EditorState,
  tr: Transaction,
  marks: readonly Mark[],
): Transaction {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return tr;
  }
  if (paragraph.textContent.length > 0) {
    return tr;
  }

  if (marks.length === 0) {
    return tr.setNodeMarkup($from.before(), undefined, {
      ...paragraph.attrs,
      defaultTextFormatting: null,
    });
  }

  const defaultTextFormatting = marksToTextFormatting(marks);

  return tr.setNodeMarkup($from.before(), undefined, {
    ...paragraph.attrs,
    defaultTextFormatting,
  });
}

// ============================================================================
// CORE MARK COMMANDS
// ============================================================================

function dispatchStoredMarks(
  state: EditorState,
  dispatch: (tr: Transaction) => void,
  marks: readonly Mark[],
): void {
  let tr = state.tr;
  tr = saveStoredMarksToParagraph(state, tr, marks);
  tr.setStoredMarks(marks);
  dispatch(tr);
}

function compactAttrs(attrs: MarkAttrs | undefined): MarkAttrs {
  if (!attrs) {
    return {};
  }

  const result: MarkAttrs = {};

  for (const [key, value] of Object.entries(attrs)) {
    if (value !== null && value !== undefined) {
      result[key] = value;
    }
  }

  return result;
}

function mergeMarkAttrs(
  markType: MarkType,
  currentMark: Mark | undefined,
  nextAttrs: MarkAttrs,
): MarkAttrs {
  const next = compactAttrs(nextAttrs);

  switch (markType.name) {
    case "fontFamily": {
      const current = currentMark
        ? fontFamilyAttrsToFormatting(expectFontFamilyMarkAttrs(currentMark))
        : undefined;
      const incoming = fontFamilyAttrsToFormatting(
        expectFontFamilyMarkAttrs(markType.create(next)),
      );
      return mergeFontFamily(current, incoming);
    }
    case "underline":
      return {
        ...compactAttrs(currentMark?.attrs),
        ...next,
      };
    default:
      return nextAttrs;
  }
}

function markRequiresAttrMerge(markType: MarkType): boolean {
  return markType.name === "fontFamily" || markType.name === "underline";
}

function createMarkWithMergedAttrs(
  markType: MarkType,
  currentMark: Mark | undefined,
  nextAttrs: MarkAttrs,
): Mark {
  if (!markRequiresAttrMerge(markType)) {
    return markType.create(nextAttrs);
  }

  return markType.create(mergeMarkAttrs(markType, currentMark, nextAttrs));
}

export function setMark(markType: MarkType, attrs: MarkAttrs): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;

    if (empty) {
      if (dispatch) {
        const current = state.storedMarks ?? state.selection.$from.marks();
        const currentMark = markType.isInSet(current);
        const marks = markType.isInSet(current)
          ? current.filter((m) => m.type !== markType)
          : current;
        const mark = createMarkWithMergedAttrs(markType, currentMark, attrs);

        dispatchStoredMarks(state, dispatch, [...marks, mark]);
      }
      return true;
    }

    if (dispatch) {
      if (!markRequiresAttrMerge(markType)) {
        dispatch(state.tr.addMark(from, to, markType.create(attrs)).scrollIntoView());
        return true;
      }

      let tr = state.tr;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) {
          return;
        }

        const start = Math.max(from, pos);
        const end = Math.min(to, pos + node.nodeSize);
        const currentMark = markType.isInSet(node.marks);
        const mark = createMarkWithMergedAttrs(markType, currentMark, attrs);
        tr = tr.addMark(start, end, mark);
      });

      dispatch(tr.scrollIntoView());
    }

    return true;
  };
}

type PairedToggleProperty = "bold" | "italic";
const DIRECT_FONT_PROPERTIES = ["color", "fontFamily", "fontSize"] as const;

const updatePairedToggleAttrs = (
  attrs: RunFormattingOverrideAttrs,
  property: PairedToggleProperty,
  enabled: boolean,
): RunFormattingOverrideAttrs => {
  const next = { ...attrs };
  const complexProperty = property === "bold" ? "boldCs" : "italicCs";
  if (property === "bold") {
    next.bold = enabled;
    next.boldCs = enabled;
  } else {
    next.italic = enabled;
    next.italicCs = enabled;
  }
  const absences = next.complexScriptPropertyAbsences?.filter(
    (candidate) => candidate !== complexProperty,
  );
  if (absences && absences.length > 0) {
    next.complexScriptPropertyAbsences = absences;
  } else {
    delete next.complexScriptPropertyAbsences;
  }
  return next;
};

const updateFontSizeCompanionAttrs = (
  attrs: RunFormattingOverrideAttrs,
  size: number | undefined,
): RunFormattingOverrideAttrs => {
  const next = { ...attrs };
  const directFontProperties = new Set(next.directFontProperties ?? []);
  if (size === undefined) {
    delete next.fontSizeCs;
    directFontProperties.delete("fontSize");
  } else {
    next.fontSizeCs = size;
    directFontProperties.add("fontSize");
  }
  if (directFontProperties.size > 0) {
    next.directFontProperties = DIRECT_FONT_PROPERTIES.filter((property) =>
      directFontProperties.has(property),
    );
  } else {
    delete next.directFontProperties;
  }
  const absences = next.complexScriptPropertyAbsences?.filter(
    (candidate) => candidate !== "fontSizeCs",
  );
  if (absences && absences.length > 0) {
    next.complexScriptPropertyAbsences = absences;
  } else {
    delete next.complexScriptPropertyAbsences;
  }
  return next;
};

type UpdateRunFormattingOverride = (
  attrs: RunFormattingOverrideAttrs,
) => RunFormattingOverrideAttrs;

const updateRunFormattingOverrideMarks = (
  marks: readonly Mark[],
  overrideType: MarkType,
  update: UpdateRunFormattingOverride,
): readonly Mark[] => {
  const existing = overrideType.isInSet(marks);
  const attrs = update(existing ? expectRunFormattingOverrideMarkAttrs(existing) : {});
  const withoutExisting = marks.filter((mark) => mark.type !== overrideType);
  return hasRunFormattingOverrideAttrs(attrs)
    ? overrideType.create(attrs).addToSet(withoutExisting)
    : withoutExisting;
};

const updateRunFormattingOverride = (
  state: EditorState,
  tr: Transaction,
  update: UpdateRunFormattingOverride,
): Transaction => {
  const overrideType = state.schema.marks["runFormattingOverride"];
  if (!overrideType) {
    return tr;
  }
  const { from, to, empty } = state.selection;
  if (empty) {
    const marks = updateRunFormattingOverrideMarks(
      tr.storedMarks ?? state.storedMarks ?? state.selection.$from.marks(),
      overrideType,
      update,
    );
    saveStoredMarksToParagraph(state, tr, marks);
    tr.setStoredMarks(marks);
    return tr;
  }

  const representations = selectRunFormattingCarrierRepresentations({ doc: tr.doc, from, to });
  for (const representation of representations) {
    const { node, position } = representation;
    const nextMarks = updateRunFormattingOverrideMarks(node.marks, overrideType, update);
    if (node.isText) {
      tr.removeMark(representation.from, representation.to, overrideType);
      const nextOverride = overrideType.isInSet(nextMarks);
      if (nextOverride) {
        tr.addMark(representation.from, representation.to, nextOverride);
      }
      continue;
    }
    tr.setNodeMarkup(position, undefined, node.attrs, nextMarks);
  }
  return tr;
};

const withRunFormattingOverride =
  (command: Command, update: UpdateRunFormattingOverride): Command =>
  (state, dispatch) =>
    command(
      state,
      dispatch
        ? (tr) => {
            dispatch(updateRunFormattingOverride(state, tr, update));
          }
        : undefined,
    );

/** UI toggle whose direct-formatting contract explicitly targets both script families. */
export const toggleMarkForAllScripts =
  (markType: MarkType, property: PairedToggleProperty): Command =>
  (state, dispatch) => {
    const enabled = !isMarkActive(state, markType);
    return withRunFormattingOverride(toggleMark(markType), (attrs) =>
      updatePairedToggleAttrs(attrs, property, enabled),
    )(state, dispatch);
  };

/** UI size command whose direct-formatting contract explicitly targets both script families. */
export const setFontSizeForAllScripts = (markType: MarkType, size: number): Command =>
  withRunFormattingOverride(setMark(markType, { size }), (attrs) =>
    updateFontSizeCompanionAttrs(attrs, size),
  );

/** Clears both ordinary and complex-script direct size through the UI command boundary. */
export const clearFontSizeForAllScripts = (markType: MarkType): Command =>
  withRunFormattingOverride(removeMark(markType), (attrs) =>
    updateFontSizeCompanionAttrs(attrs, undefined),
  );

function selectionHasVisibleUnderline(state: EditorState, markType: MarkType): boolean {
  const { from, to, empty, $from } = state.selection;

  if (empty) {
    const mark = markType.isInSet(state.storedMarks ?? $from.marks());
    return mark !== undefined && mark.attrs["style"] !== "none";
  }

  let hasVisibleUnderline = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (!node.isText) {
      return true;
    }

    const mark = markType.isInSet(node.marks);
    if (mark && mark.attrs["style"] !== "none") {
      hasVisibleUnderline = true;
      return false;
    }

    return true;
  });

  return hasVisibleUnderline;
}

export function toggleUnderlineMark(markType: MarkType): Command {
  return (state, dispatch) =>
    setMark(markType, {
      style: selectionHasVisibleUnderline(state, markType) ? "none" : "single",
    })(state, dispatch);
}

export function removeMark(markType: MarkType): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;

    if (empty) {
      if (dispatch) {
        const marks = (state.storedMarks ?? state.selection.$from.marks()).filter(
          (m) => m.type !== markType,
        );
        dispatchStoredMarks(state, dispatch, marks);
      }
      return true;
    }

    if (dispatch) {
      dispatch(state.tr.removeMark(from, to, markType).scrollIntoView());
    }

    return true;
  };
}

/**
 * Check if a mark is active in the current selection
 */
export function isMarkActive(
  state: EditorState,
  markType: MarkType,
  attrs?: Record<string, unknown>,
): boolean {
  const { from, to, empty } = state.selection;

  if (empty) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    return marks.some((mark) => {
      if (mark.type !== markType) {
        return false;
      }
      if (!attrs) {
        return true;
      }
      return Object.entries(attrs).every(([key, value]) => mark.attrs[key] === value);
    });
  }

  let hasMark = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isText) {
      const mark = markType.isInSet(node.marks);
      if (mark) {
        if (!attrs) {
          hasMark = true;
          return false;
        }
        const attrsMatch = Object.entries(attrs).every(([key, value]) => mark.attrs[key] === value);
        if (attrsMatch) {
          hasMark = true;
          return false;
        }
      }
    }
    return true;
  });

  return hasMark;
}

/**
 * Get the current value of a mark attribute
 */
export function getMarkAttr(state: EditorState, markType: MarkType, attr: string): unknown {
  const { empty, $from, from, to } = state.selection;

  if (empty) {
    const marks = state.storedMarks ?? $from.marks();
    for (const mark of marks) {
      if (mark.type === markType) {
        return mark.attrs[attr];
      }
    }
    return null;
  }

  let value: unknown = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isText && value === null) {
      const mark = markType.isInSet(node.marks);
      if (mark) {
        value = mark.attrs[attr];
        return false;
      }
    }
    return true;
  });

  return value;
}

const addDirectFontProvenance = (
  marks: Mark[],
  schema: Schema,
  directFormatting: TextFormatting | undefined,
): void => {
  const directFontProperties: ("fontFamily" | "fontSize" | "color")[] = [];
  if (directFormatting?.fontFamily !== undefined) {
    directFontProperties.push("fontFamily");
  }
  if (directFormatting?.fontSize !== undefined) {
    directFontProperties.push("fontSize");
  }
  if (directFormatting?.color !== undefined) {
    directFontProperties.push("color");
  }
  if (directFontProperties.length === 0) {
    return;
  }

  const index = marks.findIndex(({ type }) => type.name === "runFormattingOverride");
  const existing = index >= 0 ? marks.at(index) : undefined;
  const override = schema.mark("runFormattingOverride", {
    ...existing?.attrs,
    directFontProperties,
  });
  if (index >= 0) {
    marks[index] = override;
    return;
  }
  marks.push(override);
};

const COMPLEX_SCRIPT_MIRRORS = [
  { ordinary: "bold", complex: "boldCs" },
  { ordinary: "italic", complex: "italicCs" },
  { ordinary: "fontSize", complex: "fontSizeCs" },
] as const satisfies readonly {
  ordinary: keyof TextFormatting;
  complex: ComplexScriptRunPropertyKey;
}[];

const addComplexScriptAbsenceProvenance = (
  marks: Mark[],
  schema: Schema,
  directFormatting: TextFormatting | undefined,
): void => {
  const absent = COMPLEX_SCRIPT_MIRRORS.filter(
    ({ ordinary, complex }) =>
      directFormatting?.[ordinary] !== undefined && directFormatting[complex] === undefined,
  ).map(({ complex }) => complex);
  if (absent.length === 0) {
    return;
  }

  const index = marks.findIndex(({ type }) => type.name === "runFormattingOverride");
  const existing = index >= 0 ? marks.at(index) : undefined;
  const existingAbsences = new Set(existing?.attrs["complexScriptPropertyAbsences"] ?? []);
  for (const property of absent) {
    existingAbsences.add(property);
  }
  const override = schema.mark("runFormattingOverride", {
    ...existing?.attrs,
    complexScriptPropertyAbsences: COMPLEX_SCRIPT_RUN_PROPERTY_KEYS.filter((property) =>
      existingAbsences.has(property),
    ),
  });
  if (index >= 0) {
    marks[index] = override;
    return;
  }
  marks.push(override);
};

export type AuthoredRunFormattingCarrier = "preserve" | "reconstruct";

/**
 * Convert TextFormatting to ProseMirror marks
 */
type TextFormattingToMarksOptions = {
  overrideFormatting: TextFormatting | undefined;
  /** Direct standard font properties whose authored provenance must survive. */
  directFormatting?: TextFormatting | undefined;
  /** Whether direct authorship is reconstructible from structural marks and the style context. */
  authoredCarrier?: AuthoredRunFormattingCarrier;
};

export function textFormattingToMarks(
  formatting: TextFormatting | undefined,
  schema: Schema,
  options?: TextFormattingToMarksOptions,
): Mark[] {
  if (!formatting) {
    return [];
  }

  const marks: Mark[] = [];
  const overrideFormatting = options ? options.overrideFormatting : formatting;
  let overrideAttrs: ReturnType<typeof buildRunFormattingOverrideAttrs>;
  if (options?.authoredCarrier === "reconstruct") {
    overrideAttrs = buildRunFormattingOverrideAttrs(overrideFormatting, {
      type: "structural-only",
    });
  } else if (options) {
    overrideAttrs = buildRunFormattingOverrideAttrs(overrideFormatting, {
      type: "authored-baseline",
      formatting: options.directFormatting,
    });
  } else {
    overrideAttrs = buildRunFormattingOverrideAttrs(overrideFormatting);
  }

  if (overrideAttrs) {
    marks.push(schema.mark("runFormattingOverride", overrideAttrs));
  }

  // Bold
  if (formatting.bold) {
    marks.push(schema.mark("bold"));
  }

  // Italic
  if (formatting.italic) {
    marks.push(schema.mark("italic"));
  }

  // Underline
  if (formatting.underline && formatting.underline.style !== "none") {
    marks.push(
      schema.mark("underline", {
        style: formatting.underline.style,
        color: formatting.underline.color,
      }),
    );
  }

  // Strikethrough
  if (formatting.strike || formatting.doubleStrike) {
    marks.push(
      schema.mark("strike", {
        double: formatting.doubleStrike || false,
      }),
    );
  }

  // Text color
  if (formatting.color && !formatting.color.auto) {
    marks.push(
      schema.mark("textColor", {
        rgb: formatting.color.rgb,
        themeColor: formatting.color.themeColor,
        themeTint: formatting.color.themeTint,
        themeShade: formatting.color.themeShade,
      }),
    );
  }

  // Highlight
  if (formatting.highlight && formatting.highlight !== "none") {
    marks.push(
      schema.mark("highlight", {
        color: formatting.highlight,
      }),
    );
  }

  // Run shading (w:shd) used as a run background. Folio models highlight as a
  // strict OOXML named-palette union, so an arbitrary fill (e.g. a Word/Google
  // Docs run background) round-trips as a dedicated runShading mark instead of
  // silently disappearing at PM conversion. eigenpal #722 (#712).
  const runShadingMarkAttrs = shadingToRunShadingAttrs(formatting.shading);
  if (runShadingMarkAttrs) {
    marks.push(schema.mark("runShading", runShadingMarkAttrs));
  }

  // Font size
  if (formatting.fontSize) {
    marks.push(
      schema.mark("fontSize", {
        size: formatting.fontSize,
      }),
    );
  }

  // Font family
  if (formatting.fontFamily) {
    marks.push(
      schema.mark("fontFamily", {
        ascii: formatting.fontFamily.ascii,
        hAnsi: formatting.fontFamily.hAnsi,
        eastAsia: formatting.fontFamily.eastAsia,
        cs: formatting.fontFamily.cs,
        hint: formatting.fontFamily.hint,
        asciiTheme: formatting.fontFamily.asciiTheme,
        hAnsiTheme: formatting.fontFamily.hAnsiTheme,
        eastAsiaTheme: formatting.fontFamily.eastAsiaTheme,
        csTheme: formatting.fontFamily.csTheme,
      }),
    );
  }

  if (formatting.language) {
    marks.push(schema.mark("language", formatting.language));
  }

  // Superscript/Subscript
  if (formatting.vertAlign === "superscript") {
    marks.push(schema.mark("superscript"));
  } else if (formatting.vertAlign === "subscript") {
    marks.push(schema.mark("subscript"));
  }

  // All caps (w:caps)
  if (formatting.allCaps) {
    marks.push(schema.mark("allCaps"));
  }

  // Small caps (w:smallCaps)
  if (formatting.smallCaps) {
    marks.push(schema.mark("smallCaps"));
  }

  // Character spacing (spacing, position, scale, kerning)
  const spacing = typeof formatting.spacing === "number" ? formatting.spacing : null;
  const position = typeof formatting.position === "number" ? formatting.position : null;
  const scale = normalizeHorizontalScalePercent(formatting.scale) ?? null;
  const kerning = typeof formatting.kerning === "number" ? formatting.kerning : null;
  if (spacing !== null || position !== null || scale !== null || kerning !== null) {
    marks.push(
      schema.mark("characterSpacing", {
        spacing,
        position,
        scale,
        kerning,
      }),
    );
  }

  // Hidden text (w:vanish). eigenpal #424 (gap 9).
  if (formatting.hidden === true) {
    marks.push(schema.mark("hidden"));
  }

  // Emboss (w:emboss)
  if (formatting.emboss) {
    marks.push(schema.mark("emboss"));
  }

  // Imprint/Engrave (w:imprint)
  if (formatting.imprint) {
    marks.push(schema.mark("imprint"));
  }

  // Text shadow (w:shadow)
  if (formatting.shadow) {
    marks.push(schema.mark("textShadow"));
  }

  // Emphasis mark (w:em)
  if (formatting.emphasisMark && formatting.emphasisMark !== "none") {
    marks.push(schema.mark("emphasisMark", { type: formatting.emphasisMark }));
  }

  // Text outline (w:outline)
  if (formatting.outline) {
    marks.push(schema.mark("textOutline"));
  }

  // eigenpal #424 (gap 10) — per-run RTL direction (w:rtl)
  if (formatting.rtl) {
    marks.push(schema.mark("rtl"));
  }

  // eigenpal #424 (gap 11) — text effect animation (w:effect)
  if (formatting.effect && formatting.effect !== "none") {
    marks.push(schema.mark("textEffect", { effect: formatting.effect }));
  }

  addDirectFontProvenance(marks, schema, options?.directFormatting);
  if (options?.authoredCarrier === "preserve") {
    addComplexScriptAbsenceProvenance(marks, schema, options.directFormatting);
  }

  return marks;
}

/**
 * Clear all text formatting (remove all marks)
 */
export const clearFormatting: Command = (state, dispatch) => {
  const { from, to, empty } = state.selection;

  if (empty) {
    if (dispatch) {
      // Clear the paragraph's run defaults too, so EmptyParagraphFormatExtension
      // doesn't re-derive stored marks from them right after the clear.
      const tr = saveStoredMarksToParagraph(state, state.tr, []);
      tr.setStoredMarks([]);
      dispatch(tr);
    }
    return true;
  }

  if (dispatch) {
    let tr = state.tr;

    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.isText && node.marks.length > 0) {
        const start = Math.max(from, pos);
        const end = Math.min(to, pos + node.nodeSize);
        for (const mark of node.marks) {
          tr = tr.removeMark(start, end, mark.type);
        }
      }
    });

    dispatch(tr.scrollIntoView());
  }

  return true;
};

/**
 * Create a command that sets a mark on the selection
 */
export function createSetMarkCommand(markType: MarkType, attrs?: Record<string, unknown>): Command {
  return setMark(markType, attrs ?? {});
}

/**
 * Create a command that removes a mark from the selection
 */
export function createRemoveMarkCommand(markType: MarkType): Command {
  return removeMark(markType);
}
