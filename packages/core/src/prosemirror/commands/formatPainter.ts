/**
 * Format Painter commands — copy the character formatting of one selection and
 * apply ("paint") it onto another, mirroring a word processor's copy/paste
 * formatting (Ctrl+Shift+C / Ctrl+Shift+V).
 *
 * These are pure, headless commands: capture reads the marks at the current
 * selection, apply lays a captured mark set onto the current selection's range.
 * The armed/sticky toolbar interaction lives in the React layer; all editor-state
 * logic is here so it can be unit-tested without a DOM.
 */

import type { Mark, Node as PMNode, ResolvedPos } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

import { expectCharacterStyleMarkAttrs } from "../attrs";
import { getDocumentStyleResolver } from "../plugins/documentStyles";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "../runFormattingReconciliation";
import {
  paragraphFormattingForRun,
  paragraphRunStyleContext,
  paragraphRunStyleContextAt,
  resolveEffectiveRunStyleFormatting,
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "../runStyleFormatting";
import type { RunFormattingOverrideAttrs } from "../schema/marks";
import type { TextFormatting } from "../../types/document";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
/**
 * Character-formatting marks the painter copies. Structural marks — comments,
 * hyperlinks, tracked changes (insertion/deletion), footnote references — are
 * deliberately excluded so painting never moves an anchor, link, or revision.
 * `rtl` (run direction) and `hidden` are excluded too: painting them could flip
 * text direction or hide content, which is surprising for a formatting brush.
 */
export const PAINTABLE_MARK_NAMES: ReadonlySet<string> = new Set([
  "bold",
  "italic",
  "underline",
  "strike",
  "textColor",
  "highlight",
  "fontSize",
  "fontFamily",
  "superscript",
  "subscript",
  "allCaps",
  "smallCaps",
  "characterSpacing",
  "runShading",
  "runFormattingOverride",
  "emboss",
  "imprint",
  "textShadow",
  "emphasisMark",
  "textOutline",
  "textEffect",
  "characterStyle",
]);

/**
 * Formatting that must never be painted — whether it arrives as its own mark or
 * is smuggled in as a `runFormattingOverride` attr (e.g. `rtl: false`,
 * `hidden: false`). Single source of truth: it gates both the mark-type
 * exclusion and the override-attr filtering below, so an exclusion can't leak
 * back in through the override mark as new override attrs are added.
 */
const NON_PAINTABLE_FORMATS: ReadonlySet<string> = new Set(["rtl", "hidden"]);

/** A mark is paintable when it is in the allowlist and not a hard exclusion. */
function isPaintableMark(mark: Mark): boolean {
  const name = mark.type.name;
  return PAINTABLE_MARK_NAMES.has(name) && !NON_PAINTABLE_FORMATS.has(name);
}

/**
 * Strip excluded overrides (`rtl`/`hidden`) from a `runFormattingOverride` mark
 * so they cannot be re-smuggled past the mark-type exclusion. Returns the mark
 * unchanged when it carries no excluded attr, a trimmed mark when it carries a
 * mix, or null when only excluded overrides remain (nothing paintable left).
 */
function sanitizeOverrideMark(mark: Mark): Mark | null {
  const kept: Record<string, unknown> = {};
  let removedExcluded = false;

  for (const [key, value] of Object.entries(mark.attrs)) {
    if (value === null || value === undefined) {
      continue;
    }
    if (NON_PAINTABLE_FORMATS.has(key)) {
      removedExcluded = true;
      continue;
    }
    kept[key] = value;
  }

  for (const provenanceKey of ["_authoredOn", "_authoredOff"] as const) {
    const properties = kept[provenanceKey];
    if (!Array.isArray(properties)) {
      continue;
    }
    const paintableProperties = properties.filter(
      (property) => typeof property === "string" && !NON_PAINTABLE_FORMATS.has(property),
    );
    if (paintableProperties.length !== properties.length) {
      kept[provenanceKey] = paintableProperties;
      removedExcluded = true;
    }
  }

  if (!removedExcluded) {
    return mark;
  }
  if (Object.keys(kept).length === 0) {
    return null;
  }
  return mark.type.create(kept);
}

type PaintableFormatting = Pick<
  TextFormatting,
  | "allCaps"
  | "bold"
  | "boldCs"
  | "color"
  | "cs"
  | "doubleStrike"
  | "effect"
  | "emboss"
  | "emphasisMark"
  | "fontFamily"
  | "fontSize"
  | "fontSizeCs"
  | "highlight"
  | "imprint"
  | "italic"
  | "italicCs"
  | "kerning"
  | "outline"
  | "position"
  | "scale"
  | "shading"
  | "shadow"
  | "smallCaps"
  | "spacing"
  | "strike"
  | "underline"
  | "vertAlign"
>;

type BooleanOverrideKey = keyof Pick<
  RunFormattingOverrideAttrs,
  | "allCaps"
  | "bold"
  | "boldCs"
  | "cs"
  | "doubleStrike"
  | "emboss"
  | "imprint"
  | "italic"
  | "italicCs"
  | "outline"
  | "shadow"
  | "smallCaps"
  | "strike"
>;

type SentinelOverrideKey = keyof Pick<
  RunFormattingOverrideAttrs,
  | "color"
  | "effect"
  | "emphasisMark"
  | "highlight"
  | "kerning"
  | "position"
  | "scale"
  | "shading"
  | "spacing"
  | "underline"
  | "vertAlign"
>;

type SentinelTargetRelativeFormattingPolicy = {
  [K in SentinelOverrideKey]: {
    type: "sentinel";
    override: K;
    inactive: boolean | number | string;
    cancellation: NonNullable<PaintableFormatting[K]>;
    projection?: "shading" | "style";
  };
}[SentinelOverrideKey];

type TargetRelativeFormattingPolicy =
  | {
      type: "boolean";
      override: BooleanOverrideKey;
      fallback?: keyof PaintableFormatting;
    }
  | SentinelTargetRelativeFormattingPolicy
  | { type: "source-value-required"; fallback?: keyof PaintableFormatting };

/**
 * Every effective property represented by a paintable mark has one target-context
 * policy. Adding a new paintable formatting field therefore requires an explicit
 * decision: an OOXML off sentinel, a neutral scalar, or a source value that cannot
 * be synthesized without flattening document defaults.
 */
const TARGET_RELATIVE_FORMATTING_POLICIES = {
  allCaps: { type: "boolean", override: "allCaps" },
  bold: { type: "boolean", override: "bold" },
  boldCs: {
    type: "boolean",
    override: "boldCs",
    fallback: "bold",
  },
  color: {
    type: "sentinel",
    override: "color",
    inactive: "auto",
    cancellation: { auto: true },
  },
  cs: { type: "boolean", override: "cs" },
  doubleStrike: { type: "boolean", override: "doubleStrike" },
  effect: { type: "sentinel", override: "effect", inactive: "none", cancellation: "none" },
  emboss: { type: "boolean", override: "emboss" },
  emphasisMark: {
    type: "sentinel",
    override: "emphasisMark",
    inactive: "none",
    cancellation: "none",
  },
  fontFamily: { type: "source-value-required" },
  fontSize: { type: "source-value-required" },
  fontSizeCs: { type: "source-value-required", fallback: "fontSize" },
  highlight: {
    type: "sentinel",
    override: "highlight",
    inactive: "none",
    cancellation: "none",
  },
  imprint: { type: "boolean", override: "imprint" },
  italic: { type: "boolean", override: "italic" },
  italicCs: {
    type: "boolean",
    override: "italicCs",
    fallback: "italic",
  },
  kerning: { type: "sentinel", override: "kerning", inactive: 0, cancellation: 0 },
  outline: { type: "boolean", override: "outline" },
  position: { type: "sentinel", override: "position", inactive: 0, cancellation: 0 },
  scale: { type: "sentinel", override: "scale", inactive: 100, cancellation: 100 },
  shading: {
    type: "sentinel",
    override: "shading",
    inactive: "nil",
    cancellation: { pattern: "nil" },
    projection: "shading",
  },
  shadow: { type: "boolean", override: "shadow" },
  smallCaps: { type: "boolean", override: "smallCaps" },
  spacing: { type: "sentinel", override: "spacing", inactive: 0, cancellation: 0 },
  strike: { type: "boolean", override: "strike" },
  underline: {
    type: "sentinel",
    override: "underline",
    inactive: "none",
    cancellation: { style: "none" },
    projection: "style",
  },
  vertAlign: {
    type: "sentinel",
    override: "vertAlign",
    inactive: "baseline",
    cancellation: "baseline",
  },
} as const satisfies Record<keyof PaintableFormatting, TargetRelativeFormattingPolicy>;

type PaintableFormattingKey = keyof typeof TARGET_RELATIVE_FORMATTING_POLICIES;

// SAFETY: this map is the source of truth for PaintableFormattingKey.
const PAINTABLE_FORMATTING_KEYS = Object.keys(
  TARGET_RELATIVE_FORMATTING_POLICIES,
) as PaintableFormattingKey[];

const targetRelativeFormattingPolicy = (
  key: PaintableFormattingKey,
): TargetRelativeFormattingPolicy => TARGET_RELATIVE_FORMATTING_POLICIES[key];

export type CapturedTextFormatting = Readonly<{
  effectiveFormatting: Readonly<PaintableFormatting>;
  marks: readonly Mark[];
  type: "capturedTextFormatting";
}>;

function pickPaintableFormatting(formatting: TextFormatting | undefined): PaintableFormatting {
  const result: PaintableFormatting = {};
  if (!formatting) {
    return result;
  }
  for (const key of PAINTABLE_FORMATTING_KEYS) {
    const value = formatting[key];
    if (value !== undefined) {
      Reflect.set(result, key, value);
    }
  }
  return result;
}

type ResolveEffectiveFormattingOptions = {
  context: ParagraphRunStyleContext;
  marks: readonly Mark[];
  styleResolver: RunStyleResolver | null;
};

function resolveEffectiveFormatting({
  context,
  marks,
  styleResolver,
}: ResolveEffectiveFormattingOptions): PaintableFormatting {
  const authoredFormatting = readAuthoredRunFormatting({ context, marks, styleResolver });
  const paragraphFormatting = paragraphFormattingForRun(marks, context, authoredFormatting);
  const inheritedFormatting = resolveEffectiveRunStyleFormatting({
    marks,
    paragraphFormatting,
    styleResolver,
  });
  const effectiveFormatting = mergeTextFormatting(inheritedFormatting, authoredFormatting) ?? {};
  // These authored sentinels disable the inherited property as a whole. Generic
  // object merging would otherwise retain inactive decoration metadata from the
  // source style (for example, a fill beside w:shd w:val="nil").
  if (authoredFormatting.shading?.pattern === "nil") {
    effectiveFormatting.shading = authoredFormatting.shading;
  }
  if (authoredFormatting.underline?.style === "none") {
    effectiveFormatting.underline = authoredFormatting.underline;
  }
  return pickPaintableFormatting(effectiveFormatting);
}

function resolveCopiedStyleFormatting({
  context,
  marks,
  styleResolver,
}: ResolveEffectiveFormattingOptions): PaintableFormatting | null {
  const hasCharacterStyle = marks.some((mark) => mark.type.name === "characterStyle");
  if (hasCharacterStyle && !styleResolver) {
    return null;
  }
  const paragraphFormatting = paragraphFormattingForRun(marks, context);
  return pickPaintableFormatting(
    resolveEffectiveRunStyleFormatting({ marks, paragraphFormatting, styleResolver }),
  );
}

function sameFormattingValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) {
    return true;
  }
  if (left === null || right === null || typeof left !== "object" || typeof right !== "object") {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    return (
      Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((value, index) => sameFormattingValue(value, right.at(index)))
    );
  }
  const leftEntries = Object.entries(left);
  const rightKeys = Object.keys(right);
  return (
    leftEntries.length === rightKeys.length &&
    leftEntries.every(
      ([property, value]) =>
        Object.hasOwn(right, property) && sameFormattingValue(value, Reflect.get(right, property)),
    )
  );
}

function policyValue(
  formatting: Readonly<PaintableFormatting>,
  key: PaintableFormattingKey,
  policy: TargetRelativeFormattingPolicy,
): unknown {
  const own = formatting[key];
  const value = own ?? (policy.type !== "sentinel" ? formatting[policy.fallback ?? key] : own);
  if (policy.type !== "sentinel") {
    return value;
  }
  if (policy.projection === "style") {
    return typeof value === "object" && value !== null ? Reflect.get(value, "style") : undefined;
  }
  if (policy.projection === "shading") {
    if (typeof value !== "object" || value === null) {
      return undefined;
    }
    return Reflect.get(value, "pattern") === "nil" ? "nil" : value;
  }
  if (key === "color") {
    return typeof value === "object" && value !== null && Reflect.get(value, "auto") === true
      ? "auto"
      : value;
  }
  return value;
}

type MaterializeTargetRelativeMarksOptions = {
  captured: CapturedTextFormatting;
  context: ParagraphRunStyleContext;
  node: PMNode;
  styleResolver: RunStyleResolver | null;
};

function materializeTargetRelativeMarks({
  captured,
  context,
  node,
  styleResolver,
}: MaterializeTargetRelativeMarksOptions): readonly Mark[] | null {
  const targetStyleFormatting = resolveCopiedStyleFormatting({
    context,
    marks: captured.marks,
    styleResolver,
  });
  if (!targetStyleFormatting) {
    return null;
  }
  const targetRelativeFormatting: TextFormatting = {};

  for (const key of PAINTABLE_FORMATTING_KEYS) {
    const policy = targetRelativeFormattingPolicy(key);
    const sourceValue = policyValue(captured.effectiveFormatting, key, policy);
    const targetValue = policyValue(targetStyleFormatting, key, policy);

    // An omitted complex-script companion follows its ordinary property. Keep
    // that authored absence when neither style defines an independent value;
    // writing the fallback as direct formatting would invent provenance.
    if (
      (key === "boldCs" || key === "italicCs" || key === "fontSizeCs") &&
      captured.effectiveFormatting[key] === undefined &&
      targetStyleFormatting[key] === undefined
    ) {
      continue;
    }

    if (policy.type === "source-value-required") {
      if (sourceValue === undefined && targetValue !== undefined) {
        return null;
      }
      if (sourceValue !== undefined && !sameFormattingValue(sourceValue, targetValue)) {
        Reflect.set(targetRelativeFormatting, key, sourceValue);
      }
      continue;
    }
    if (policy.type === "boolean") {
      const sourceIsOn = sourceValue === true;
      if (sourceIsOn !== (targetValue === true)) {
        Reflect.set(targetRelativeFormatting, policy.override, sourceIsOn);
      }
      continue;
    }

    const sourceIsInactive = sourceValue === undefined || sourceValue === policy.inactive;
    const targetIsInactive = targetValue === undefined || targetValue === policy.inactive;
    if (sourceIsInactive) {
      if (!targetIsInactive) {
        Reflect.set(
          targetRelativeFormatting,
          policy.override,
          captured.effectiveFormatting[key] ?? policy.cancellation,
        );
      }
      continue;
    }
    if (!sameFormattingValue(sourceValue, targetValue)) {
      const formattingValue = captured.effectiveFormatting[key];
      if (formattingValue === undefined) {
        return null;
      }
      Reflect.set(targetRelativeFormatting, policy.override, formattingValue);
    }
  }

  const characterStyle = captured.marks.find((mark) => mark.type.name === "characterStyle");
  const characterStyleId = characterStyle
    ? expectCharacterStyleMarkAttrs(characterStyle).styleId
    : undefined;
  const marks = reconcileRunFormattingMarks({
    authoredFormatting: {
      ...targetRelativeFormatting,
      ...(characterStyleId !== undefined ? { styleId: characterStyleId } : {}),
    },
    context,
    node,
    styleResolver,
  });
  return marks.filter(isPaintableMark);
}

/**
 * Marks of the first text node inside [from, to). A word processor's format
 * brush copies from the start of the source selection, so a mixed selection
 * yields the formatting of its leading run.
 */
type FirstTextFormatting = {
  marks: readonly Mark[];
  position: ResolvedPos;
};

function firstTextFormatting(
  state: EditorState,
  from: number,
  to: number,
): FirstTextFormatting | null {
  let formatting: FirstTextFormatting | null = null;

  state.doc.nodesBetween(from, to, (node, position) => {
    if (formatting) {
      return false;
    }
    if (node.isText) {
      formatting = { marks: node.marks, position: state.doc.resolve(position) };
      return false;
    }
    return true;
  });

  return formatting;
}

/**
 * Capture the paintable character formatting active over the current selection.
 * An empty selection reads the stored/insertion marks; a range reads its leading
 * run. The discriminated result deliberately represents an effectively plain run:
 * `null` belongs to the host's “nothing captured” state, not this command.
 */
export function captureFormatMarks(state: EditorState): CapturedTextFormatting {
  const { empty, $from, from, to } = state.selection;
  const firstText = empty ? null : firstTextFormatting(state, from, to);
  const source = empty ? (state.storedMarks ?? $from.marks()) : (firstText?.marks ?? []);

  const captured: Mark[] = [];
  for (const mark of source) {
    if (!isPaintableMark(mark)) {
      continue;
    }
    if (mark.type.name === "runFormattingOverride") {
      const sanitized = sanitizeOverrideMark(mark);
      if (sanitized) {
        captured.push(sanitized);
      }
      continue;
    }
    captured.push(mark);
  }
  const styleResolver = getDocumentStyleResolver(state);
  const context = paragraphRunStyleContextAt(
    state.doc,
    (firstText?.position ?? $from).pos,
    styleResolver,
  );
  return {
    effectiveFormatting: resolveEffectiveFormatting({ context, marks: captured, styleResolver }),
    marks: captured,
    type: "capturedTextFormatting",
  };
}

type TargetParagraphSegment = {
  context: ParagraphRunStyleContext;
  from: number;
  to: number;
};

function targetParagraphSegments(
  state: EditorState,
  from: number,
  to: number,
  styleResolver: RunStyleResolver | null,
): TargetParagraphSegment[] {
  const segments: TargetParagraphSegment[] = [];
  state.doc.nodesBetween(from, to, (node, position) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const contentFrom = position + 1;
    const segmentFrom = Math.max(from, contentFrom);
    const segmentTo = Math.min(to, contentFrom + node.content.size);
    if (segmentTo > segmentFrom) {
      segments.push({
        context: paragraphRunStyleContext(node, styleResolver),
        from: segmentFrom,
        to: segmentTo,
      });
    }
    return false;
  });
  return segments;
}

/**
 * Apply captured marks onto the current selection's range: clear every paintable
 * mark type first (replace, not merge — so painting Georgia-12-bold onto
 * Arial-10 leaves only Georgia-12-bold), then lay down the captured marks.
 *
 * Best-effort: a collapsed selection, no capture, or a target-relative delta
 * that cannot be represented without guessing document defaults is a no-op.
 * Returns false so a keymap binding stays unclaimed; never partially mutates.
 */
export function applyFormatMarks(captured: CapturedTextFormatting | null): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty || captured?.type !== "capturedTextFormatting") {
      return false;
    }

    const styleResolver = getDocumentStyleResolver(state);
    const segments = targetParagraphSegments(state, from, to, styleResolver);
    const node = state.schema.text("format-painter");
    const targetMarks = segments.map((segment) =>
      materializeTargetRelativeMarks({
        captured,
        context: segment.context,
        node,
        styleResolver,
      }),
    );
    if (segments.length === 0 || targetMarks.some((marks) => marks === null)) {
      return false;
    }

    if (!dispatch) {
      return true;
    }

    let tr = state.tr;
    for (const name of PAINTABLE_MARK_NAMES) {
      const markType = state.schema.marks[name];
      if (markType) {
        tr = tr.removeMark(from, to, markType);
      }
    }
    for (const [index, segment] of segments.entries()) {
      const marks = targetMarks[index];
      if (!marks) {
        continue;
      }
      for (const mark of marks) {
        tr = tr.addMark(segment.from, segment.to, mark);
      }
    }

    dispatch(tr.scrollIntoView());
    return true;
  };
}
