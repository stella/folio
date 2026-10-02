/**
 * `LISTNUM` fields a numbered paragraph draws as part of its list marker.
 *
 * The reader takes each such field, and the tab after it, out of the
 * paragraph's content and shows the field's cached result in the marker. The
 * two functions here are the two halves of that: {@link foldListNumberFields}
 * takes them out and records what they were and where they stood, and
 * {@link withFoldedListNumberFields} puts them back for a save.
 *
 * A position is counted in inline units rather than content items, because
 * the editor regroups text into runs of its own: two runs the field stood
 * between come back as one, and the field goes back between the same two
 * characters.
 */

import type {
  ComplexField,
  FoldedListNumberField,
  FoldedListNumberFields,
  Paragraph,
  ParagraphContent,
  Run,
} from "../types/document";

/** Zero-width markup the reader steps over between a field and its tab. */
const RANGE_MARKER_TYPES: ReadonlySet<string> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);

/** Items that occupy no inline unit and are not counted as range markers. */
const UNCOUNTED_TYPES: ReadonlySet<string> = new Set([
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "renderedPageBreak",
]);

/** Items whose inline units are those of what they hold. */
const CONTAINER_TYPES: ReadonlySet<string> = new Set([
  "run",
  "hyperlink",
  "insertion",
  "deletion",
  "moveFrom",
  "moveTo",
  "inlineSdt",
  "inlineWrapper",
]);

const isRangeMarker = (content: ParagraphContent): boolean => RANGE_MARKER_TYPES.has(content.type);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/**
 * Inline units an item occupies: one per character of text, one per other
 * inline object, and the sum of its members for an item that only groups them.
 */
const inlineUnits = (item: unknown): number => {
  if (!isRecord(item)) {
    return 0;
  }
  const type = item["type"];
  if (typeof type !== "string" || RANGE_MARKER_TYPES.has(type) || UNCOUNTED_TYPES.has(type)) {
    return 0;
  }
  if (type === "text") {
    const text = item["text"];
    return typeof text === "string" ? text.length : 0;
  }
  if (CONTAINER_TYPES.has(type)) {
    const members = type === "hyperlink" ? item["children"] : item["content"];
    if (!Array.isArray(members)) {
      return 0;
    }
    let units = 0;
    for (const member of members) {
      units += inlineUnits(member);
    }
    return units;
  }
  return 1;
};

const isListNumberField = (content: ParagraphContent): content is ComplexField =>
  content.type === "complexField" &&
  (content.fieldType === "LISTNUM" ||
    content.instruction.trim().toUpperCase().startsWith("LISTNUM"));

const isTabRun = (content: ParagraphContent): content is Run =>
  content.type === "run" && content.content.length === 1 && content.content[0]?.type === "tab";

const cachedText = (field: ComplexField): string => {
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

export type ListNumberFieldFold = {
  /** The content without the fields and the tabs that followed them. */
  content: ParagraphContent[];
  /** The cached display of each field that has one, in source order. */
  cached: string[];
  /** Every field taken out, with where it stood. */
  fields: FoldedListNumberField[];
};

/**
 * Take every `LISTNUM` field, and the tab after it, out of `content`.
 *
 * Bookmark and comment markers between a field and its tab stay in the
 * content: the tab is still the field's when only they separate the two.
 */
export const foldListNumberFields = (content: readonly ParagraphContent[]): ListNumberFieldFold => {
  const kept: ParagraphContent[] = [];
  const cached: string[] = [];
  const fields: FoldedListNumberField[] = [];
  let offset = 0;
  // Range markers kept since `offset` last grew.
  let markers = 0;
  let awaitingTab: FoldedListNumberField | undefined;
  let gap = 0;

  for (const item of content) {
    if (awaitingTab && isRangeMarker(item)) {
      kept.push(item);
      markers += 1;
      gap += 1;
      continue;
    }
    if (awaitingTab) {
      const field = awaitingTab;
      awaitingTab = undefined;
      if (isTabRun(item)) {
        field.tab = item;
        field.markersBeforeTab = gap;
        continue;
      }
    }
    if (isListNumberField(item)) {
      const field: FoldedListNumberField = { field: item, offset, markersBefore: markers };
      fields.push(field);
      const text = cachedText(item);
      if (text) {
        cached.push(text);
      }
      awaitingTab = field;
      gap = 0;
      continue;
    }
    kept.push(item);
    const units = inlineUnits(item);
    if (units > 0) {
      offset += units;
      markers = 0;
    } else if (isRangeMarker(item)) {
      markers += 1;
    }
  }

  return { content: kept, cached, fields };
};

/**
 * The fields `paragraph` still owns. They were folded into the marker of one
 * numbering level, so a paragraph that has left that level, or lost its
 * numbering, shows none of them and writes none.
 */
export const foldedListNumberFieldsOf = (
  paragraph: Pick<Paragraph, "foldedListNumberFields" | "listRendering">,
): FoldedListNumberFields | undefined => {
  const folded = paragraph.foldedListNumberFields;
  const rendering = paragraph.listRendering;
  if (
    !folded ||
    folded.fields.length === 0 ||
    !rendering ||
    rendering.numId !== folded.numId ||
    rendering.level !== folded.level
  ) {
    return undefined;
  }
  return folded;
};

/** `run` cut after `units` inline units. */
const splitRun = (run: Run, units: number): [Run, Run] => {
  const head: Run["content"] = [];
  const tail: Run["content"] = [];
  let remaining = units;
  for (const piece of run.content) {
    if (remaining <= 0) {
      tail.push(piece);
      continue;
    }
    const pieceUnits = inlineUnits(piece);
    if (pieceUnits <= remaining) {
      head.push(piece);
      remaining -= pieceUnits;
      continue;
    }
    // Only text is wider than one unit, so only text is ever cut.
    if (piece.type === "text") {
      head.push({ ...piece, text: piece.text.slice(0, remaining) });
      tail.push({ ...piece, text: piece.text.slice(remaining) });
    } else {
      tail.push(piece);
    }
    remaining = 0;
  }
  return [
    { ...run, content: head },
    { ...run, content: tail },
  ];
};

type Placement = {
  offset: number;
  /** Range markers at `offset` that come before the item. */
  markers: number;
  item: ParagraphContent;
};

/**
 * The content of `paragraph` with its folded fields and their tabs back where
 * they stood.
 *
 * A field goes back at its recorded offset, after the range markers that
 * preceded it there. Where an edit left less than was recorded, it goes as
 * far as the content reaches: ahead of the first visible item when markers
 * are missing, and at the end when the content is shorter than the offset.
 */
export const withFoldedListNumberFields = (paragraph: Paragraph): ParagraphContent[] => {
  const folded = foldedListNumberFieldsOf(paragraph);
  if (!folded) {
    return paragraph.content;
  }

  const placements: Placement[] = [];
  for (const { field, tab, offset, markersBefore, markersBeforeTab } of folded.fields) {
    placements.push({ offset, markers: markersBefore, item: field });
    if (tab) {
      placements.push({ offset, markers: markersBefore + (markersBeforeTab ?? 0), item: tab });
    }
  }

  const restored: ParagraphContent[] = [];
  let next = 0;
  let offset = 0;
  let markers = 0;

  /** Write every placement that belongs ahead of `upcoming`. */
  const place = (upcoming: ParagraphContent | undefined): void => {
    for (let placement = placements[next]; placement; placement = placements[next]) {
      const due =
        upcoming === undefined ||
        placement.offset < offset ||
        (placement.offset === offset &&
          (placement.markers <= markers || inlineUnits(upcoming) > 0));
      if (!due) {
        return;
      }
      restored.push(placement.item);
      next += 1;
    }
  };

  for (const item of paragraph.content) {
    let rest: ParagraphContent | undefined = item;
    while (rest) {
      place(rest);
      const units = inlineUnits(rest);
      const placement = placements[next];
      if (
        placement &&
        rest.type === "run" &&
        placement.offset > offset &&
        placement.offset < offset + units
      ) {
        const [head, tail]: [Run, Run] = splitRun(rest, placement.offset - offset);
        restored.push(head);
        offset = placement.offset;
        markers = 0;
        rest = tail;
        continue;
      }
      restored.push(rest);
      if (units > 0) {
        offset += units;
        markers = 0;
      } else if (isRangeMarker(rest)) {
        markers += 1;
      }
      rest = undefined;
    }
  }
  place(undefined);

  return restored;
};
