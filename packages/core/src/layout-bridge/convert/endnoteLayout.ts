/**
 * Endnote areas (ECMA-376 §17.11).
 *
 * Endnotes are not reserved per page the way footnotes are. `w:pos` puts them
 * after the last body block of the document (`docEnd`, the default) or of each
 * section (`sectEnd`), where they paginate as body content does: they are
 * spliced into the body's block list and the same measure and paginate path
 * places them. Each area opens with the `w:separator` endnote story, and every
 * page it continues onto opens with the `w:continuationSeparator` story. The
 * parsed model keeps only normal notes, so both stories are the ones a package
 * is written with: one paragraph, no space after, single line spacing, holding
 * the separator mark.
 *
 * A note's blocks carry positions in the note's own document, so they are
 * stripped of them before joining the body flow: a click, a caret or a
 * selection in the body must never resolve into a note. The painted note is
 * addressed by its story instead, which is what the note editor opens.
 */

import { panic } from "better-result";

import { calculateColumnWidths } from "../../layout-engine/paginator";
import { resolveEffectiveParagraphSpacingTree } from "../../layout-engine/paragraphSpacing";
import {
  stripFlowBlockPmAnchors,
  type FlowBlock,
  type Measure,
  type NoteAreaLayout,
  type NoteSeparatorKind,
  type ParagraphBlock,
  type ParagraphMeasure,
} from "../../layout-engine/types";
import type {
  BlockContent,
  Endnote,
  EndnotePosition,
  SectionProperties,
} from "../../types/document";
import type { NoteStoryKey } from "../../types/editor-story";
import { getMargins, getPageSize } from "../../paged-layout/sectionGeometry";
import { getColumns } from "../sectionColumns";
import {
  collectEndnoteRefs,
  convertFootnoteToContent,
  convertNoteStoryToFlowBlocks,
  type ConvertFootnoteOptions,
  type MeasureBlocksFn,
} from "./footnoteLayout";

/** One endnote area and the body position it follows. */
type EndnoteArea = {
  /**
   * The body `sectionBreak` block the area precedes, counted from zero; `null`
   * places it after the last body block.
   */
  beforeSectionBreak: number | null;
  blocks: FlowBlock[];
  measures: Measure[];
};

export type PreparedEndnoteAreas = {
  areas: readonly EndnoteArea[];
  noteAreas: NoteAreaLayout;
  /** The endnote each area content block belongs to, by block id. */
  noteStoryByBlockId: ReadonlyMap<string, NoteStoryKey>;
};

export type PrepareEndnoteAreasInput = {
  /** The body's top-level flow blocks, markers already renumbered. */
  blocks: readonly FlowBlock[];
  endnotes: readonly Endnote[] | undefined;
  /** Endnote `w:id` → reference-order number, as the body markers show it. */
  displayNumbers: ReadonlyMap<number, number>;
  /** Endnote `w:id` → the formatted number the body marker shows. */
  displayTexts: ReadonlyMap<number, string>;
  position: EndnotePosition;
  /**
   * Section properties by section index: `w:noEndnote`, and the column width
   * an area wraps to in the section it follows.
   */
  sections: readonly (SectionProperties | undefined)[];
  options: ConvertFootnoteOptions & { measureBlocks: MeasureBlocksFn };
};

/**
 * The separator story a notes part is written with (folio writes it too):
 * `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`
 * around the mark run, in the document's default paragraph style.
 */
const SEPARATOR_STORY: BlockContent[] = [
  {
    type: "paragraph",
    formatting: { spaceAfter: 0, lineSpacing: 240, lineSpacingRule: "auto" },
    content: [],
  },
];

/**
 * Convert and measure every referenced endnote into the areas `w:pos` asks
 * for. `undefined` when no reference names a note, so a document without
 * endnotes lays out exactly as it did before.
 */
export function prepareEndnoteAreas(
  input: PrepareEndnoteAreasInput,
): PreparedEndnoteAreas | undefined {
  const notesById = new Map<number, Endnote>();
  for (const note of input.endnotes ?? []) {
    if (note.noteType === "normal") {
      notesById.set(note.id, note);
    }
  }
  if (notesById.size === 0 || input.displayNumbers.size === 0) {
    return undefined;
  }

  const finalSection = countSectionBreaks(input.blocks);
  const referenceSection = firstReferenceSections(input.blocks);
  // `displayNumbers` is in reference order, which is the order notes appear in.
  const notesBySection = new Map<number, number[]>();
  for (const noteId of input.displayNumbers.keys()) {
    if (!notesById.has(noteId)) {
      continue;
    }
    const section =
      input.position === "sectEnd"
        ? collectingSection(referenceSection.get(noteId) ?? finalSection, finalSection, input)
        : finalSection;
    const noteIds = notesBySection.get(section) ?? [];
    noteIds.push(noteId);
    notesBySection.set(section, noteIds);
  }

  const separatorBlockIds = new Map<string, NoteSeparatorKind>();
  const contentBlockIds = new Set<string>();
  const noteStoryByBlockId = new Map<string, NoteStoryKey>();
  const areas: EndnoteArea[] = [];

  for (const [section, noteIds] of [...notesBySection].sort(([left], [right]) => left - right)) {
    const width = endnoteAreaWidth(
      section === finalSection ? input.sections.at(-1) : input.sections[section],
    );
    const areaKey = `endnote-area-${String(section)}`;
    const separator = convertSeparatorStory(`${areaKey}-separator`, width, input.options);
    // The separator never ends a page on its own: it travels with the first note.
    const blocks: FlowBlock[] = [
      { ...separator.block, attrs: { ...separator.block.attrs, keepNext: true } },
    ];
    const measures: Measure[] = [separator.measure];
    separatorBlockIds.set(String(separator.block.id), "separator");
    for (const noteId of noteIds) {
      const note = notesById.get(noteId)!; // SAFETY: filtered above
      const content = convertFootnoteToContent(note, input.displayNumbers.get(noteId) ?? 0, width, {
        ...input.options,
        displayText: input.displayTexts.get(noteId) ?? String(noteId),
      });
      for (const [index, block] of content.blocks.entries()) {
        const id = `endnote-${String(noteId)}-${String(index)}`;
        blocks.push(withId(stripFlowBlockPmAnchors(block), id));
        measures.push(content.measures[index]!); // SAFETY: one measure per content block
        contentBlockIds.add(id);
        noteStoryByBlockId.set(id, { kind: "endnote", noteId });
      }
    }
    areas.push({
      beforeSectionBreak: section === finalSection ? null : section,
      // Contextual spacing resolved once, as the body's blocks are, so the
      // painter reads the spacing the paginator placed.
      blocks: resolveEffectiveParagraphSpacingTree(blocks),
      measures,
    });
  }

  const continuationSeparator = convertSeparatorStory(
    "endnote-continuation-separator",
    endnoteAreaWidth(input.sections.at(-1)),
    input.options,
  );
  separatorBlockIds.set(String(continuationSeparator.block.id), "continuationSeparator");

  return {
    areas,
    noteAreas: { separatorBlockIds, contentBlockIds, continuationSeparator },
    noteStoryByBlockId,
  };
}

/**
 * The body blocks with each endnote area in place: before the section break
 * that ends the section it follows, or after the last body block.
 */
export function spliceEndnoteAreas(
  blocks: FlowBlock[],
  measures: Measure[],
  prepared: PreparedEndnoteAreas | undefined,
): { blocks: FlowBlock[]; measures: Measure[] } {
  if (prepared === undefined || prepared.areas.length === 0) {
    return { blocks, measures };
  }
  const areasByBreak = new Map<number, EndnoteArea>();
  let documentEndArea: EndnoteArea | undefined;
  for (const area of prepared.areas) {
    if (area.beforeSectionBreak === null) {
      documentEndArea = area;
    } else {
      areasByBreak.set(area.beforeSectionBreak, area);
    }
  }

  const splicedBlocks: FlowBlock[] = [];
  const splicedMeasures: Measure[] = [];
  let sectionBreakIndex = 0;
  for (const [index, block] of blocks.entries()) {
    if (block.kind === "sectionBreak") {
      const area = areasByBreak.get(sectionBreakIndex);
      if (area !== undefined) {
        splicedBlocks.push(...area.blocks);
        splicedMeasures.push(...area.measures);
      }
      sectionBreakIndex += 1;
    }
    splicedBlocks.push(block);
    splicedMeasures.push(measures[index]!); // SAFETY: one measure per body block
  }
  if (documentEndArea !== undefined) {
    splicedBlocks.push(...documentEndArea.blocks);
    splicedMeasures.push(...documentEndArea.measures);
  }
  return { blocks: splicedBlocks, measures: splicedMeasures };
}

/**
 * The endnotes a `docEnd` area collects after the body's last block, for
 * `toFlowBlocks`' `trailingEndnoteIds`. Empty for `sectEnd`, whose final area
 * depends on where the references fall.
 */
export function trailingEndnoteIds(
  endnotes: readonly Endnote[] | undefined,
  position: EndnotePosition,
): ReadonlySet<number> {
  if (position !== "docEnd") {
    return new Set();
  }
  return new Set((endnotes ?? []).flatMap((note) => (note.noteType === "normal" ? [note.id] : [])));
}

/** `w:pos` for the document: the settings value, else the final section's. */
export function resolveEndnotePosition(
  settingsPosition: EndnotePosition | undefined,
  finalSection: SectionProperties | undefined,
): EndnotePosition {
  return settingsPosition ?? finalSection?.endnotePr?.position ?? "docEnd";
}

/** The first column's width in a section: what an endnote area there wraps to. */
export function endnoteAreaWidth(section: SectionProperties | undefined): number {
  const pageSize = getPageSize(section);
  const margins = getMargins(section);
  const [firstColumn] = calculateColumnWidths(
    pageSize.w,
    margins.left,
    margins.right,
    getColumns(section) ?? { count: 1, gap: 0 },
  );
  return firstColumn ?? pageSize.w - margins.left - margins.right;
}

function countSectionBreaks(blocks: readonly FlowBlock[]): number {
  let count = 0;
  for (const block of blocks) {
    if (block.kind === "sectionBreak") {
      count += 1;
    }
  }
  return count;
}

/** The section each endnote is first referenced in. */
function firstReferenceSections(blocks: readonly FlowBlock[]): Map<number, number> {
  const sections = new Map<number, number>();
  let section = 0;
  for (const block of blocks) {
    if (block.kind === "sectionBreak") {
      section += 1;
      continue;
    }
    for (const { endnoteId } of collectEndnoteRefs([block])) {
      if (!sections.has(endnoteId)) {
        sections.set(endnoteId, section);
      }
    }
  }
  return sections;
}

/**
 * The section whose end collects a note referenced in `section`: a section
 * with `w:noEndnote` passes its endnotes on to the next section that shows
 * them, and the document's end collects whatever is left.
 */
function collectingSection(
  section: number,
  finalSection: number,
  input: PrepareEndnoteAreasInput,
): number {
  let target = section;
  while (target < finalSection && input.sections[target]?.noEndnote === true) {
    target += 1;
  }
  return target;
}

/** A separator story's paragraph, converted and measured as a note's text is. */
function convertSeparatorStory(
  id: string,
  width: number,
  options: ConvertFootnoteOptions & { measureBlocks: MeasureBlocksFn },
): { block: ParagraphBlock; measure: ParagraphMeasure } {
  const converted = convertNoteStoryToFlowBlocks(SEPARATOR_STORY, options).flowBlocks.at(0);
  const block: ParagraphBlock =
    converted?.kind === "paragraph"
      ? { ...converted, id }
      : { kind: "paragraph", id, runs: [{ kind: "text", text: "" }] };
  const [measure] = options.measureBlocks([block], width);
  if (measure?.kind !== "paragraph") {
    panic("An endnote separator paragraph must measure as a paragraph");
  }
  const { pmStart: _pmStart, pmEnd: _pmEnd, ...detached } = block;
  return { block: detached, measure };
}

function withId(block: FlowBlock, id: string): FlowBlock {
  return { ...block, id };
}
