/**
 * `LISTNUM` fields a numbered paragraph draws as part of its list marker.
 *
 * A field that opens a numbered paragraph is shown by the marker, which
 * carries its cached result, so the field and the tab after it must not show
 * on the line as well. They are not taken out of the paragraph: each stays in
 * the content, where it stood, as a capture of the markup it was read from.
 * A capture shows nothing and is written back as it stands.
 *
 * One rule decides which captures a paragraph may hold, and the reader, the
 * editor and the save all apply it ({@link planListNumberFold}): a capture is
 * hidden only while it stands at the start of a paragraph whose marker shows
 * it. A field anywhere else is on the line, as the field it is.
 */

import type {
  ComplexField,
  Deletion,
  FoldedListNumber,
  Insertion,
  MoveFrom,
  MoveTo,
  Paragraph,
  ParagraphContent,
  Run,
} from "../types/document";

/** Zero-width markup that can stand between a field and its tab. */
const RANGE_MARKER_TYPES: ReadonlySet<ParagraphContent["type"]> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);

/** Markup that shows nothing and so does not end the start of a paragraph. */
const ZERO_WIDTH_TYPES: ReadonlySet<ParagraphContent["type"]> = new Set([
  ...RANGE_MARKER_TYPES,
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

/** Whether `content` is markup that may stand between a field and the tab that follows it. */
export const isListNumberGapMarker = (content: ParagraphContent): boolean =>
  RANGE_MARKER_TYPES.has(content.type);

const isZeroWidth = (content: ParagraphContent): boolean =>
  ZERO_WIDTH_TYPES.has(content.type) || (content.type === "preservedInline" && content.text === "");

/** Whether `content` is a `LISTNUM` field, by its parsed type or its instruction. */
export const isListNumberField = (content: ParagraphContent): content is ComplexField =>
  content.type === "complexField" &&
  (content.fieldType === "LISTNUM" ||
    content.instruction.trim().toUpperCase().startsWith("LISTNUM"));

/** A run that holds one tab and nothing else. */
export const isTabOnlyRun = (content: ParagraphContent): content is Run =>
  content.type === "run" && content.content.length === 1 && content.content[0]?.type === "tab";

/** The text a field's cached result puts in the marker. */
const cachedListNumberText = (field: ComplexField): string => {
  let text = "";
  for (const run of field.fieldResult) {
    for (const piece of run.content) {
      if (piece.type === "text") {
        text += piece.text;
      }
    }
  }
  return text;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/** Whether `value` is what a capture stands for: a complex field, or a run. */
export const isFoldedListNumber = (value: unknown): value is FoldedListNumber => {
  if (!isRecord(value)) {
    return false;
  }
  if (value["kind"] === "field") {
    const field = value["field"];
    return (
      isRecord(field) &&
      field["type"] === "complexField" &&
      typeof field["instruction"] === "string" &&
      Array.isArray(field["fieldCode"]) &&
      Array.isArray(field["fieldResult"])
    );
  }
  if (value["kind"] === "tab") {
    const run = value["run"];
    return isRecord(run) && run["type"] === "run" && Array.isArray(run["content"]);
  }
  return false;
};

/** What `content` stands for when it is a capture the fold left, else nothing. */
export const foldedListNumberOf = (content: ParagraphContent): FoldedListNumber | undefined =>
  content.type === "preservedInline" ? content.foldedListNumber : undefined;

/** Whether `content` is a capture the fold left in place of a field or its tab. */
export const isFoldedListNumberCapture = (content: ParagraphContent): boolean =>
  foldedListNumberOf(content) !== undefined;

export type ListNumberFieldFold = {
  /** The content, each folded field and its tab replaced by a capture. */
  content: ParagraphContent[];
  /** The cached display of each folded field that has one, in source order. */
  cached: string[];
  /** Every `LISTNUM` field the content holds, folded or not. */
  fieldCount: number;
};

/**
 * Replace the `LISTNUM` fields that open a paragraph, and the tab after each,
 * by captures of the markup they were read from.
 *
 * Only what stands ahead of the first thing the line shows is folded: a field
 * after text is on the line where the text is. `sourceMarkupOf` answers the
 * markup; a field it has none for is left as the field it is, and ends the
 * fold. Bookmark and comment markers between a field and its tab stay where
 * they are, and the tab is still the field's when only they separate the two.
 */
export const foldListNumberFields = (
  content: readonly ParagraphContent[],
  sourceMarkupOf: (item: ComplexField | Run) => string | undefined,
): ListNumberFieldFold => {
  const folded: ParagraphContent[] = [];
  const cached: string[] = [];
  let fieldCount = 0;
  let leading = true;
  let awaitingTab = false;

  for (const item of content) {
    if (isListNumberField(item)) {
      fieldCount += 1;
    }
    if (!leading) {
      folded.push(item);
      continue;
    }
    if (awaitingTab && RANGE_MARKER_TYPES.has(item.type)) {
      folded.push(item);
      continue;
    }
    if (awaitingTab) {
      awaitingTab = false;
      const tabMarkup = isTabOnlyRun(item) ? sourceMarkupOf(item) : undefined;
      if (isTabOnlyRun(item) && tabMarkup !== undefined) {
        folded.push({
          type: "preservedInline",
          xml: tabMarkup,
          text: "",
          foldedListNumber: { kind: "tab", run: item },
        });
        continue;
      }
    }
    if (isListNumberField(item)) {
      const markup = sourceMarkupOf(item);
      if (markup !== undefined) {
        folded.push({
          type: "preservedInline",
          xml: markup,
          text: "",
          foldedListNumber: { kind: "field", field: item },
        });
        const text = cachedListNumberText(item);
        if (text) {
          cached.push(text);
        }
        awaitingTab = true;
        continue;
      }
    }
    folded.push(item);
    leading = isZeroWidth(item);
  }

  return { content: folded, cached, fieldCount };
};

/** One item of a paragraph, as far as the fold needs to know it. */
export type ListNumberFoldItem =
  /** A capture of a field, with the text its cached result shows. */
  | { kind: "field"; cached: string }
  /** A capture of the tab after a field. */
  | { kind: "tab" }
  /** Markup that shows nothing. */
  | { kind: "hidden" }
  /** Anything the line shows. */
  | { kind: "shown" };

export type ListNumberFoldPlan = {
  /** The items' indices in the order they are to stand in. */
  order: number[];
  /** Indices of the captures that stay hidden; every other capture goes on the line. */
  hidden: ReadonlySet<number>;
  /** What the marker shows after its own text, or nothing. */
  suffix: string | undefined;
};

const isCapture = (item: ListNumberFoldItem | undefined): boolean =>
  item?.kind === "field" || item?.kind === "tab";

/**
 * Decide which captures of a paragraph stay hidden, and what its marker shows.
 *
 * `markerShowsFields` says whether the paragraph's marker shows folded fields
 * at all. When it does not, no capture is hidden.
 *
 * When it does, the captures it shows are the ones that open the paragraph:
 * those ahead of the first thing the line shows, each field followed at most
 * by its own tab. Text put in front of them was put in front of the body, not
 * of the marker, so if the first capture no longer opens the paragraph, it
 * and the captures that stand with it go back to the start. Every capture
 * further in is on the line.
 */
export const planListNumberFold = (
  items: readonly ListNumberFoldItem[],
  markerShowsFields: boolean,
): ListNumberFoldPlan => {
  const unmoved = items.map((_, index) => index);
  const none: ListNumberFoldPlan = { order: unmoved, hidden: new Set(), suffix: undefined };
  const first = items.findIndex((item) => isCapture(item));
  if (!markerShowsFields || first === -1) {
    return none;
  }

  let order = unmoved;
  if (items.slice(0, first).some((item) => item.kind === "shown")) {
    // The captures that stand together with the first one, markers between
    // them included, and none of the markers after the last.
    let end = first + 1;
    for (let index = first + 1; index < items.length; index += 1) {
      const item = items[index];
      if (item?.kind === "shown") {
        break;
      }
      if (isCapture(item)) {
        end = index + 1;
      }
    }
    order = [...unmoved.slice(first, end), ...unmoved.slice(0, first), ...unmoved.slice(end)];
  }

  const hidden = new Set<number>();
  const cached: string[] = [];
  let fieldAwaitsTab = false;
  for (const index of order) {
    const item = items[index];
    if (item?.kind === "hidden") {
      continue;
    }
    if (item?.kind === "field") {
      hidden.add(index);
      if (item.cached) {
        cached.push(item.cached);
      }
      fieldAwaitsTab = true;
      continue;
    }
    if (item?.kind === "tab" && fieldAwaitsTab) {
      hidden.add(index);
      fieldAwaitsTab = false;
      continue;
    }
    break;
  }

  // A marker with nothing to show hides nothing, as the reader folds nothing.
  return cached.length === 0 ? none : { order, hidden, suffix: cached.join(" ") };
};

type ListMarkerFoldState = {
  /** Whether the marker shows folded fields after its own text. */
  showsFields: boolean;
  /** The marker's own text, without what it shows of folded fields. */
  base: string;
};

type ListMarkerFoldInput = {
  marker: string | null | undefined;
  template: string | null | undefined;
  isBullet: boolean;
  numbered: boolean;
};

/**
 * What a paragraph's marker says about folded fields. It shows them after a
 * tab its own level text does not have; a bullet shows only its glyph.
 */
const listMarkerFoldState = ({
  marker,
  template,
  isBullet,
  numbered,
}: ListMarkerFoldInput): ListMarkerFoldState => {
  const tab = marker?.indexOf("\t") ?? -1;
  if (!marker || !numbered || isBullet || !template || template.includes("\t") || tab === -1) {
    return { showsFields: false, base: marker ?? "" };
  }
  return { showsFields: true, base: marker.slice(0, tab) };
};

/** The marker text for `state` showing `suffix`. */
const listMarkerWithFold = (state: ListMarkerFoldState, suffix: string | undefined): string =>
  suffix === undefined ? state.base : `${state.base}\t${suffix}`;

const foldItemOf = (content: ParagraphContent): ListNumberFoldItem => {
  const folded = foldedListNumberOf(content);
  if (folded) {
    return folded.kind === "field"
      ? { kind: "field", cached: cachedListNumberText(folded.field) }
      : { kind: "tab" };
  }
  return isZeroWidth(content) ? { kind: "hidden" } : { kind: "shown" };
};

/** The item a capture stands for, which shows on the line; any other item is itself. */
export const unfoldedListNumberContent = (content: ParagraphContent): ParagraphContent => {
  const folded = foldedListNumberOf(content);
  if (!folded) {
    return content;
  }
  return folded.kind === "field" ? folded.field : folded.run;
};

type TrackedChange = Insertion | Deletion | MoveFrom | MoveTo;

/**
 * The one definition of a tracked change for the fold: content inserted,
 * deleted, or moved from or to a place under revision tracking.
 */
const isTrackedChange = (content: ParagraphContent): content is TrackedChange =>
  content.type === "insertion" ||
  content.type === "deletion" ||
  content.type === "moveFrom" ||
  content.type === "moveTo";

/**
 * `content` with every capture inside a tracked change replaced by the item
 * it stands for. A tracked change is written with its own spellings (deleted
 * text is `w:delText`, a deleted instruction `w:delInstrText`), which markup
 * replayed as it was read does not have; the field and the run do.
 */
const withTrackedCapturesShown = (content: ParagraphContent[]): ParagraphContent[] => {
  let changed = false;
  const next = content.map((item): ParagraphContent => {
    if (!isTrackedChange(item)) {
      return item;
    }
    if (!item.content.some((child) => foldedListNumberOf(child) !== undefined)) {
      return item;
    }
    changed = true;
    return {
      ...item,
      content: item.content.map((child) => {
        const folded = foldedListNumberOf(child);
        if (!folded) {
          return child;
        }
        return structuredClone(folded.kind === "field" ? folded.field : folded.run);
      }),
    };
  });
  return changed ? next : content;
};

/**
 * Bring `paragraph` to the one form the fold allows: the captures its marker
 * shows stand at its start, every other capture is the field or the tab it
 * stood for, and the marker shows exactly the fields still hidden behind it.
 * A capture under a tracked change is never hidden.
 */
export const normalizeFoldedListNumbers = (paragraph: Paragraph): void => {
  paragraph.content = withTrackedCapturesShown(paragraph.content);
  const rendering = paragraph.listRendering;
  const state = listMarkerFoldState({
    marker: rendering?.marker,
    template: rendering?.markerTemplate,
    isBullet: rendering?.isBullet === true,
    numbered: rendering !== undefined,
  });
  const content = paragraph.content;
  if (!state.showsFields && !content.some(isFoldedListNumberCapture)) {
    return;
  }

  const plan = planListNumberFold(content.map(foldItemOf), state.showsFields);
  const next: ParagraphContent[] = [];
  for (const index of plan.order) {
    const item = content[index];
    if (item === undefined) {
      continue;
    }
    // A copy: the capture may be shared with the editor state it came from,
    // and what goes on the line is the paragraph's own to change.
    next.push(
      plan.hidden.has(index) || !isFoldedListNumberCapture(item)
        ? item
        : structuredClone(unfoldedListNumberContent(item)),
    );
  }
  paragraph.content = next;
  if (rendering && state.showsFields) {
    rendering.marker = listMarkerWithFold(state, plan.suffix);
  }
};
