/**
 * Note areas that flow with the body (ECMA-376 §17.11).
 *
 * Endnotes are collected after the last body block of the document, or of a
 * section, and paginate exactly as body content does. Two things set them
 * apart from the body: the area opens with its `w:separator` story, and every
 * page the area continues onto opens with its `w:continuationSeparator` story.
 * Both stories draw their separator mark as a horizontal rule.
 */

import { measuredLineRangeHeight } from "./lineFlow";
import type { PageTopContinuation, Paginator } from "./paginator";
import { getParagraphSpacingAfter } from "./paragraphSpacing";
import {
  NOTE_SEPARATOR_RULE_THICKNESS,
  NOTE_SEPARATOR_WIDTH_FRACTION,
  type FlowBlock,
  type Layout,
  type Measure,
  type NoteAreaLayout,
  type NoteSeparatorKind,
  type NoteSeparatorRule,
  type ParagraphFragment,
  type ParagraphMeasure,
} from "./types";

/**
 * The continuation separator a page opens with while a note area continues
 * onto it. Its space after collapses with the next note paragraph's space
 * before, as between any two paragraphs.
 */
export function createNoteAreaContinuation(
  noteAreas: NoteAreaLayout | undefined,
  paginator: Paginator,
): PageTopContinuation | undefined {
  const continuation = noteAreas?.continuationSeparator;
  if (continuation === undefined) {
    return undefined;
  }
  const { block, measure } = continuation;
  const height = measuredLineRangeHeight(measure.lines, 0, measure.lines.length);
  return (state, x, width) => {
    const fragment: ParagraphFragment = {
      kind: "paragraph",
      blockId: block.id,
      x,
      y: state.cursorY,
      width,
      height,
      fromLine: 0,
      toLine: measure.lines.length,
    };
    paginator.addUnflowedFragment(fragment);
    state.trailingSpacing = getParagraphSpacingAfter(block);
    return height;
  };
}

/** Whether laying out `block` continues a note area onto any page it opens. */
export function continuesNoteArea(
  noteAreas: NoteAreaLayout | undefined,
  block: FlowBlock,
): boolean {
  return noteAreas?.contentBlockIds.has(String(block.id)) === true;
}

/**
 * Record where each separator mark landed. A rule is centred on the first line
 * of the separator paragraph: a `w:separator` spans part of the column and a
 * `w:continuationSeparator` all of it.
 */
export function annotateNoteSeparators(
  layout: Layout,
  blocks: readonly FlowBlock[],
  measures: readonly Measure[],
  noteAreas: NoteAreaLayout | undefined,
): Layout {
  if (noteAreas === undefined || noteAreas.separatorBlockIds.size === 0) {
    return layout;
  }
  const measureById = new Map<string, ParagraphMeasure>();
  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (noteAreas.separatorBlockIds.has(String(block.id)) && measure?.kind === "paragraph") {
      measureById.set(String(block.id), measure);
    }
  }
  const continuation = noteAreas.continuationSeparator;
  if (continuation !== undefined) {
    measureById.set(String(continuation.block.id), continuation.measure);
  }

  for (const page of layout.pages) {
    const rules: NoteSeparatorRule[] = [];
    for (const fragment of page.fragments) {
      if (fragment.kind !== "paragraph" || fragment.fromLine !== 0) {
        continue;
      }
      const kind = noteAreas.separatorBlockIds.get(String(fragment.blockId));
      const firstLine = measureById.get(String(fragment.blockId))?.lines.at(0);
      if (kind === undefined || firstLine === undefined) {
        continue;
      }
      rules.push({
        x: fragment.x,
        y: fragment.y + (firstLine.lineHeight - NOTE_SEPARATOR_RULE_THICKNESS) / 2,
        width: fragment.width * separatorWidthFraction(kind),
      });
    }
    if (rules.length > 0) {
      page.noteSeparators = rules;
    }
  }
  return layout;
}

function separatorWidthFraction(kind: NoteSeparatorKind): number {
  switch (kind) {
    case "separator":
      return NOTE_SEPARATOR_WIDTH_FRACTION;
    case "continuationSeparator":
      return 1;
    default: {
      const exhaustive: never = kind;
      return exhaustive;
    }
  }
}
