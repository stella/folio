/** Shared run merge rules for parsing and paragraph joins. */
import type {
  Run,
  RunContent,
  TextFormatting,
  ExhaustiveFields,
  PreservedAttribute,
  PreservedMarkup,
} from "../model/document";

/**
 * Every `TextFormatting` field, named so the comparison below is total.
 *
 * A field this comparison forgets is a field one of two merged runs loses,
 * silently and at parse time. Naming them turns a field added to the model
 * without a comparison into a compile error rather than a fidelity defect.
 */
type ComparedTextFormattingField =
  | "bold"
  | "boldCs"
  | "italic"
  | "italicCs"
  | "underline"
  | "strike"
  | "doubleStrike"
  | "vertAlign"
  | "smallCaps"
  | "allCaps"
  | "hidden"
  | "noProof"
  | "color"
  | "highlight"
  | "shading"
  | "fontSize"
  | "fontSizeCs"
  | "fontFamily"
  | "language"
  | "spacing"
  | "position"
  | "scale"
  | "kerning"
  | "effect"
  | "emphasisMark"
  | "emboss"
  | "imprint"
  | "outline"
  | "shadow"
  | "rtl"
  | "cs"
  | "styleId"
  | "preserved";
type ComparedTextFormatting = ExhaustiveFields<TextFormatting, ComparedTextFormattingField>;

/**
 * Check if two TextFormatting objects are equivalent
 *
 * Uses deep comparison of all properties to determine if runs
 * can be merged without losing formatting information.
 */
export function formattingEquals(
  a: ComparedTextFormatting | undefined,
  b: ComparedTextFormatting | undefined,
): boolean {
  // Both undefined - equal
  if (!a && !b) {
    return true;
  }

  // One undefined - not equal
  if (!a || !b) {
    return false;
  }

  // Compare boolean properties
  if (a.bold !== b.bold) {
    return false;
  }
  if (a.boldCs !== b.boldCs) {
    return false;
  }
  if (a.italic !== b.italic) {
    return false;
  }
  if (a.italicCs !== b.italicCs) {
    return false;
  }
  if (a.strike !== b.strike) {
    return false;
  }
  if (a.doubleStrike !== b.doubleStrike) {
    return false;
  }
  if (a.smallCaps !== b.smallCaps) {
    return false;
  }
  if (a.allCaps !== b.allCaps) {
    return false;
  }
  if (a.hidden !== b.hidden) {
    return false;
  }
  if (a.noProof !== b.noProof) {
    return false;
  }
  if (a.emboss !== b.emboss) {
    return false;
  }
  if (a.imprint !== b.imprint) {
    return false;
  }
  if (a.outline !== b.outline) {
    return false;
  }
  if (a.shadow !== b.shadow) {
    return false;
  }
  if (a.rtl !== b.rtl) {
    return false;
  }
  if (a.cs !== b.cs) {
    return false;
  }

  // Compare numeric properties
  if (a.fontSize !== b.fontSize) {
    return false;
  }
  if (a.fontSizeCs !== b.fontSizeCs) {
    return false;
  }
  if (a.spacing !== b.spacing) {
    return false;
  }
  if (a.position !== b.position) {
    return false;
  }
  if (a.scale !== b.scale) {
    return false;
  }
  if (a.kerning !== b.kerning) {
    return false;
  }

  // Compare string properties
  if (a.vertAlign !== b.vertAlign) {
    return false;
  }
  if (a.highlight !== b.highlight) {
    return false;
  }
  if (a.effect !== b.effect) {
    return false;
  }
  if (a.emphasisMark !== b.emphasisMark) {
    return false;
  }
  if (a.styleId !== b.styleId) {
    return false;
  }

  // Compare underline (object with style and optional color)
  if (!underlineEquals(a.underline, b.underline)) {
    return false;
  }

  // Compare color (object with rgb, themeColor, etc.)
  if (!colorEquals(a.color, b.color)) {
    return false;
  }

  // Compare shading (object with color, fill, pattern)
  if (!shadingEquals(a.shading, b.shading)) {
    return false;
  }

  // Compare fontFamily (complex object with multiple properties)
  if (!fontFamilyEquals(a.fontFamily, b.fontFamily)) {
    return false;
  }
  if (!languageEquals(a.language, b.language)) {
    return false;
  }

  // The `w:rPr` children no reader took a value from. A merged run carries one
  // property set, so two runs whose captured bytes differ are two runs.
  if (!preservedMarkupEquals(a.preserved, b.preserved)) {
    return false;
  }

  return true;
}

/**
 * Compare two verbatim sinks position by position.
 *
 * The sink is ordered — by `index`, then by source order within an index — and
 * the order is what puts the markup back between the same modelled siblings,
 * so equality is by sequence, not by set.
 */
function preservedMarkupEquals(
  a: PreservedMarkup | undefined,
  b: PreservedMarkup | undefined,
): boolean {
  const left = a?.children ?? [];
  const right = b?.children ?? [];
  if (left.length !== right.length) {
    return false;
  }

  return left.every(({ index, xml }, position) => {
    const other = right[position];
    return other !== undefined && other.index === index && other.xml === xml;
  });
}

/**
 * Compare two attribute remainders as sets.
 *
 * An element cannot carry one expanded name twice, so (namespace, name) orders
 * a remainder totally and the sorted keys compare exactly. Order is a set
 * comparison rather than a sequence one because attribute order in XML says
 * nothing: two runs that spelled the same attributes in a different order
 * carry the same remainder and may still become one run.
 */
function preservedAttributesEqual(
  a: readonly PreservedAttribute[] | undefined,
  b: readonly PreservedAttribute[] | undefined,
): boolean {
  const left = canonicalRemainderKeys(a);
  const right = canonicalRemainderKeys(b);
  return left.length === right.length && left.every((key, index) => key === right[index]);
}

/** `\u0000` cannot appear in an XML name, namespace URI or attribute value. */
const REMAINDER_KEY_SEPARATOR = "\u0000";

const canonicalRemainderKeys = (attributes: readonly PreservedAttribute[] | undefined): string[] =>
  (attributes ?? [])
    .map(({ namespace, name, value }) =>
      [namespace ?? "", name, value].join(REMAINDER_KEY_SEPARATOR),
    )
    .sort();

function languageEquals(a: TextFormatting["language"], b: TextFormatting["language"]): boolean {
  return a?.val === b?.val && a?.eastAsia === b?.eastAsia && a?.bidi === b?.bidi;
}

/**
 * Compare underline settings
 */
function underlineEquals(a: TextFormatting["underline"], b: TextFormatting["underline"]): boolean {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }

  if (a.style !== b.style) {
    return false;
  }
  return colorEquals(a.color, b.color);
}

/**
 * Compare color values
 */
function colorEquals(a: TextFormatting["color"], b: TextFormatting["color"]): boolean {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }

  return (
    a.rgb === b.rgb &&
    a.auto === b.auto &&
    a.themeColor === b.themeColor &&
    a.themeTint === b.themeTint &&
    a.themeShade === b.themeShade
  );
}

/**
 * Compare shading properties
 */
function shadingEquals(a: TextFormatting["shading"], b: TextFormatting["shading"]): boolean {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }

  if (a.pattern !== b.pattern) {
    return false;
  }
  if (!colorEquals(a.color, b.color)) {
    return false;
  }
  if (!colorEquals(a.fill, b.fill)) {
    return false;
  }

  return true;
}

/**
 * Compare font family settings
 */
function fontFamilyEquals(
  a: TextFormatting["fontFamily"],
  b: TextFormatting["fontFamily"],
): boolean {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }

  return (
    a.ascii === b.ascii &&
    a.hAnsi === b.hAnsi &&
    a.eastAsia === b.eastAsia &&
    a.cs === b.cs &&
    a.asciiTheme === b.asciiTheme &&
    a.hAnsiTheme === b.hAnsiTheme &&
    a.eastAsiaTheme === b.eastAsiaTheme &&
    a.csTheme === b.csTheme
  );
}

/**
 * Check if a run contains only text content
 * (runs with special content like images, fields, etc. should not be merged)
 */
export function isTextOnlyRun(run: Run): boolean {
  return run.content.every(
    (c) => c.type === "text" || c.type === "softHyphen" || c.type === "noBreakHyphen",
  );
}

/**
 * Check if run content can be merged (simple text types)
 */
function isMergeableContent(content: RunContent): boolean {
  return (
    content.type === "text" || content.type === "softHyphen" || content.type === "noBreakHyphen"
  );
}

/**
 * Check if a run can be merged with another run
 * Runs with breaks, tabs, images, fields, etc. act as merge boundaries
 */
function canMergeRun(run: Run): boolean {
  // A run-property revision owns an exact text range. Merging either boundary
  // would discard that range because a consolidated run can carry only one
  // property-change collection.
  if (run.propertyChanges && run.propertyChanges.length > 0) {
    return false;
  }

  // A run holding no payload never reaches a merge: `consolidateRuns` flushes
  // at it and keeps it whole, because the payload a later pass will put in it
  // is not this pass's to merge away.
  if (run.content.length === 0) {
    return false;
  }

  // Runs with only text/hyphen content can be merged
  return run.content.every(isMergeableContent);
}

/**
 * Every `Run` field, named by what a merge does with it.
 *
 * The merged run is the survivor spread whole, so a field added to the model
 * is carried over from one side by default. Naming each one here forces the
 * next field to state whether that default is right before it can compile.
 */
type MergeDecidedRunField =
  /** Identical for every run. */
  | "type"
  /** Concatenated; see {@link mergeRunContent}. */
  | "content"
  /** Must be equal, and the survivor keeps it. */
  | "formatting"
  /** Must be equal, and the survivor keeps it. */
  | "preservedAttributes"
  /** Refuses the merge outright; see {@link canMergeRun}. */
  | "propertyChanges";
type MergeDecidedRun = ExhaustiveFields<Run, MergeDecidedRunField>;

/**
 * May two adjacent runs become one?
 *
 * Every consolidation site asks this one question, because a merge that any
 * one site decides differently is a merge the next parse undoes. A run holds
 * three records a merged run can hold only one of: its typed formatting, the
 * attributes `w:r` carried that the model has no field for, and the `w:rPr`
 * children no reader took a value from. Merging two runs that disagree on any
 * of them discards the loser's copy, which is how both runs' `w:rsid*` used to
 * vanish at parse time — before the editor, before any gate could see it.
 */
export function runsMergeable(a: MergeDecidedRun, b: MergeDecidedRun): boolean {
  return (
    canMergeRun(a) &&
    canMergeRun(b) &&
    formattingEquals(a.formatting, b.formatting) &&
    preservedAttributesEqual(a.preservedAttributes, b.preservedAttributes)
  );
}

/**
 * Merge the content of two runs into a single content array
 */
export function mergeRunContent(content1: RunContent[], content2: RunContent[]): RunContent[] {
  // Combine all content
  const result: RunContent[] = [];

  // Add all from first run
  for (const c of content1) {
    result.push(c);
  }

  // Merge text at boundary if possible
  const lastText = result.at(-1);
  const firstText = content2[0];
  if (lastText?.type === "text" && firstText?.type === "text") {
    // SAFETY: lastText is the same reference as result.at(-1), which
    // returned a value, so result is non-empty and length - 1 is a valid index.
    result[result.length - 1] = {
      type: "text",
      text: lastText.text + firstText.text,
    };

    // Add rest of content2
    for (let i = 1; i < content2.length; i++) {
      // SAFETY: i < content2.length in for loop
      result.push(content2[i]!);
    }
  } else {
    // Just append all of content2
    for (const c of content2) {
      result.push(c);
    }
  }

  return result;
}
