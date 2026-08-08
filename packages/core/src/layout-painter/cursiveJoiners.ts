/**
 * Repair cursive letter connections that a face change severed.
 *
 * Browsers shape across an inline box boundary only while no shaping-relevant
 * property changes. Colour and underline are safe, so tracked-change and
 * comment marks already join. A face change is not: bold, italic, a different
 * size or a different family selects another font, shaping stops there, and a
 * cursive word split mid-word by such a run falls back to isolated forms and
 * visibly comes apart. Word joins straight through the same boundary.
 *
 * The repair is a zero-width joiner on each side of the boundary, carried in its
 * OWN span rather than appended to the run's text. That distinction is
 * load-bearing: `data-pm-start`/`data-pm-end` spans map DOM text offsets back to
 * ProseMirror positions for hit-testing, so growing a run's text node by a
 * character would desync every offset after it. A joiner span carries no pm
 * attributes, so the offset contract is untouched.
 *
 * Whether two runs share a face is decided by running the painter's own
 * `applyRunStyles` over a probe element for each side and diffing the result,
 * rather than by re-deriving the face from run fields. There is no second copy
 * of the font decisions to drift from the first, and a property this module has
 * never heard of counts as a face change: over-inserting a joiner is invisible,
 * missing one is not.
 *
 * KNOWN RESIDUAL — the measurer does not see this repair. Canvas `measureText`
 * ignores a joiner at a string edge (measuring the two halves with and without
 * one returns bit-identical widths), while DOM layout applies the joined forms.
 * So a repaired word paints at a different width than the measurer reserved, by
 * roughly the difference between isolated and medial advances.
 *
 * Which DIRECTION depends on the font: measured against two Arabic fallback
 * faces the divergence was -2.7px on one and +2.5px on the other. So it is a
 * bounded disagreement that can push a line either way, NOT a safe
 * over-reservation, and a line can therefore come out marginally too long. It
 * arises only on words that contain a face change, and
 * `tests/visual/measure-parity.spec.ts` holds it to a budget. Closing it needs widths from a real shaper
 * rather than canvas; until then the trade is a correct rendering against a
 * slightly wrong measurement, which beats a measurement that exactly matches a
 * visibly broken rendering.
 */

import { hasCursiveLetter, joinsAcrossBoundary } from "../utils/cursiveJoining";

/** U+200D ZERO WIDTH JOINER — Joining_Type C, so it joins on both sides. */
const ZERO_WIDTH_JOINER = "‍";

/** Marks a span as joiner furniture: no text of its own, no pm positions. */
export const JOINER_DATASET_KEY = "docxJoiner";

/**
 * Inline properties that `applyRunStyles` may set which provably cannot change
 * how text shapes, and so must NOT be read as a face change.
 *
 * A deny-list rather than an allow-list on purpose: an unlisted new property is
 * treated as face-changing, which inserts a harmless extra joiner instead of
 * silently stopping the repair. Fails toward the visible-correct rendering.
 */
const SHAPING_NEUTRAL_STYLE_PROPERTIES: ReadonlySet<string> = new Set([
  "color",
  "--doc-run-color",
  "backgroundColor",
  "background-color",
  "textDecoration",
  "text-decoration",
  "textDecorationColor",
  "text-decoration-color",
  "textDecorationStyle",
  "text-decoration-style",
  "textUnderlineOffset",
  "text-underline-offset",
  "verticalAlign",
  "vertical-align",
  "opacity",
  "borderBottom",
  "border-bottom",
  "outline",
]);

/** `applyRunStyles`, injected so this module does not import the painter back. */
export type ApplyRunStyles<TRun> = (element: HTMLElement, run: TRun) => void;

/**
 * The style declaration a run paints with, as a plain comparable record.
 *
 * Reading the keys the painter actually assigned (rather than enumerating a
 * CSSStyleDeclaration) keeps this working against both the real DOM and the
 * minimal fake the painter tests use.
 */
function paintedStyle<TRun>(
  run: TRun,
  applyRunStyles: ApplyRunStyles<TRun>,
  doc: Document,
): Record<string, string> {
  const probe = doc.createElement("span");
  applyRunStyles(probe, run);
  const style: Record<string, string> = {};
  for (const [property, value] of styleEntries(probe.style)) {
    if (value === "") continue;
    if (SHAPING_NEUTRAL_STYLE_PROPERTIES.has(property)) continue;
    style[property] = value;
  }
  return style;
}

/**
 * Enumerate the properties a declaration actually has set.
 *
 * A real `CSSStyleDeclaration` is not a plain object: `Object.entries` on one
 * yields its INDEXED entries, so `["0", "font-family"]` rather than
 * `["fontFamily", "Arial"]`, which compares the list of property names instead
 * of their values. The painter's test fake IS a plain object. Both shapes have
 * to work, and the difference between them is invisible in unit tests, so it is
 * handled here explicitly rather than left to whichever one a caller happens to
 * pass. Real declarations report kebab-case names, including custom properties.
 */
function styleEntries(style: CSSStyleDeclaration): [string, string][] {
  const declaration = style as unknown as {
    length?: number;
    item?: (index: number) => string;
    getPropertyValue?: (property: string) => string;
  };
  if (
    typeof declaration.length === "number" &&
    typeof declaration.item === "function" &&
    typeof declaration.getPropertyValue === "function"
  ) {
    const entries: [string, string][] = [];
    for (let index = 0; index < declaration.length; index++) {
      const property = declaration.item(index);
      entries.push([property, declaration.getPropertyValue(property)]);
    }
    return entries;
  }
  return Object.entries(style as unknown as Record<string, unknown>).filter(
    (entry): entry is [string, string] => typeof entry[1] === "string",
  );
}

const sameFace = (a: Record<string, string>, b: Record<string, string>): boolean => {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const key of keys) {
    if (a[key] !== b[key]) return false;
  }
  return true;
};

/** Which side(s) of a run need a joiner span. */
export type JoinerSides = { leading: boolean; trailing: boolean };

type PlanOptions<TRun> = {
  /** Runs in paint order, already sliced for the line and script-split. */
  runs: readonly TRun[];
  /** Text of a run, or undefined for a run that paints no text (image, break). */
  textOf: (run: TRun) => string | undefined;
  applyRunStyles: ApplyRunStyles<TRun>;
  doc: Document;
};

/**
 * Decide, for each run, whether it needs a joiner on either side.
 *
 * Only adjacent text-bearing runs are considered: an image or a tab between two
 * words means they were never connected, so no joiner belongs there.
 */
export function planCursiveJoiners<TRun>({
  runs,
  textOf,
  applyRunStyles,
  doc,
}: PlanOptions<TRun>): Map<TRun, JoinerSides> {
  const plan = new Map<TRun, JoinerSides>();
  if (runs.length < 2) {
    return plan;
  }

  // Gate on the whole line first: an all-Latin line must not pay for probe
  // elements or style diffing, which is the overwhelmingly common case.
  let lineHasCursive = false;
  for (const run of runs) {
    const text = textOf(run);
    if (text !== undefined && hasCursiveLetter(text)) {
      lineHasCursive = true;
      break;
    }
  }
  if (!lineHasCursive) {
    return plan;
  }

  const styleCache = new Map<TRun, Record<string, string>>();
  const styleOf = (run: TRun): Record<string, string> => {
    const cached = styleCache.get(run);
    if (cached) return cached;
    const computed = paintedStyle(run, applyRunStyles, doc);
    styleCache.set(run, computed);
    return computed;
  };

  const mark = (run: TRun, side: keyof JoinerSides): void => {
    const existing = plan.get(run);
    if (existing) {
      existing[side] = true;
      return;
    }
    plan.set(run, { leading: side === "leading", trailing: side === "trailing" });
  };

  for (let i = 0; i < runs.length - 1; i++) {
    // SAFETY: both indices are inside the loop bounds.
    const before = runs[i]!;
    const after = runs[i + 1]!;
    const beforeText = textOf(before);
    const afterText = textOf(after);
    if (beforeText === undefined || afterText === undefined) continue;
    if (!joinsAcrossBoundary(beforeText, afterText)) continue;
    if (sameFace(styleOf(before), styleOf(after))) continue;
    mark(before, "trailing");
    mark(after, "leading");
  }

  return plan;
}

/**
 * Build a joiner span wearing `run`'s own painted styles, so shaping crosses
 * from that run into the joiner and takes the joined form.
 *
 * Without the matching styles the joiner would sit in a different face and the
 * repair would be a no-op that still looked applied.
 */
export function createJoinerSpan<TRun>(
  run: TRun,
  applyRunStyles: ApplyRunStyles<TRun>,
  doc: Document,
): HTMLElement {
  const span = doc.createElement("span");
  applyRunStyles(span, run);
  span.dataset[JOINER_DATASET_KEY] = "true";
  span.textContent = ZERO_WIDTH_JOINER;
  return span;
}

/**
 * A run element wrapped with whatever joiners its boundaries need, ready to
 * spread into `append`. Returns the element alone when no repair applies.
 */
export function withCursiveJoiners<TRun>(
  runEl: HTMLElement,
  run: TRun,
  plan: Map<TRun, JoinerSides>,
  applyRunStyles: ApplyRunStyles<TRun>,
  doc: Document,
): HTMLElement[] {
  const sides = plan.get(run);
  if (!sides) {
    return [runEl];
  }
  const elements: HTMLElement[] = [];
  if (sides.leading) elements.push(createJoinerSpan(run, applyRunStyles, doc));
  elements.push(runEl);
  if (sides.trailing) elements.push(createJoinerSpan(run, applyRunStyles, doc));
  return elements;
}
