/**
 * Section Break Commands
 * @packageDocumentation
 * @public
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";

import type { SectionProperties } from "../../types/document";
import {
  mintSectionProperties,
  sectionPropertiesOf,
  type SectionBreakType,
} from "../sectionCarrier";

type InsertableSectionBreak = Extract<SectionBreakType, "nextPage" | "continuous">;

/** The position of a paragraph's mark: the end of the content it closes. */
type SectionEndingMark = {
  markPosition: number;
  properties: SectionProperties;
};

const sectionEndingMarks = (doc: PMNode): SectionEndingMark[] => {
  const marks: SectionEndingMark[] = [];
  doc.descendants((node, position) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const properties = sectionPropertiesOf(node);
    if (properties) {
      marks.push({ markPosition: position + node.nodeSize - 1, properties });
    }
    return false;
  });
  return marks;
};

const withSectionProperties = (
  attrs: PMNode["attrs"],
  properties: SectionProperties | null,
): Record<string, unknown> => {
  const next = { ...attrs };
  if (properties === null) {
    delete next["_sectionProperties"];
    return next;
  }
  next["_sectionProperties"] = properties;
  return next;
};

/**
 * Restate, over the whole document, which paragraph marks end a section.
 *
 * The rule is total in both directions, because a join is: it consumes one
 * paragraph mark and keeps the other, so exactly the marks that survived the
 * transaction still end the sections they ended before. ProseMirror keeps the
 * *first* node's attrs when it joins, which gets both halves of the rule wrong
 * on its own — the mark that survived a backward join loses a break it still
 * owns, and the node that absorbed a section-ending predecessor keeps a break
 * whose mark is gone.
 *
 * Word deletes the section break in the second case: with the caret at the
 * start of the paragraph *after* a break, Backspace removes the break, and the
 * paragraphs it governed join the following section, whose `w:sectPr` then
 * governs them (ECMA-376 Part 1 §17.6.18: a paragraph's `w:sectPr` states the
 * properties of the section ending at that paragraph, so content with no
 * `w:sectPr` before it belongs to the next one).
 */
const reconcileSectionEndingMarks = (
  transaction: Transaction,
  marks: readonly SectionEndingMark[],
): void => {
  const survivors = new Map<number, SectionProperties>();
  for (const { markPosition, properties } of marks) {
    const mapped = transaction.mapping.mapResult(markPosition, -1);
    if (!mapped.deleted) {
      survivors.set(mapped.pos, properties);
    }
  }

  const corrections: { position: number; properties: SectionProperties | null }[] = [];
  transaction.doc.descendants((node, position) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const survived = survivors.get(position + node.nodeSize - 1) ?? null;
    if (survived !== sectionPropertiesOf(node)) {
      corrections.push({ position, properties: survived });
    }
    return false;
  });

  // An attr-only markup keeps every node's size, so the positions collected
  // above stay valid as the corrections are applied one after another.
  for (const { position, properties } of corrections) {
    const paragraph = transaction.doc.nodeAt(position);
    if (paragraph) {
      transaction.setNodeMarkup(
        position,
        undefined,
        withSectionProperties(paragraph.attrs, properties),
      );
    }
  }
};

/**
 * Run a command, keeping every section break on the paragraph whose mark
 * survived it.
 *
 * A section break is a property of a paragraph *mark*. The transfer is keyed on
 * the mark's position surviving the transaction, not on which command ran, so
 * deleting a section-ending paragraph outright still removes its section: that
 * mark is gone, and nothing inherits it.
 */
export const keepSectionBreaksOnSurvivingMarks =
  (command: Command): Command =>
  (state, dispatch, view) => {
    if (!dispatch) {
      return command(state, undefined, view);
    }
    const marks = sectionEndingMarks(state.doc);
    if (marks.length === 0) {
      return command(state, dispatch, view);
    }
    let captured: Transaction | null = null;
    const capture = (transaction: Transaction): void => {
      captured = transaction;
    };
    if (!command(state, capture, view)) {
      return false;
    }
    // Commands that report success without dispatching leave the document as
    // it was, so there is no mark to follow.
    const transaction: Transaction | null = captured;
    if (transaction === null) {
      return true;
    }
    reconcileSectionEndingMarks(transaction, marks);
    dispatch(transaction);
    return true;
  };

/**
 * What the section record of a paragraph in the selection becomes. Called once
 * per record the selection holds, and once per paragraph that holds none.
 */
type SectionRecordRewrite = (current: SectionProperties | null) => SectionProperties | null;

const selectedParagraphs = (state: EditorState): { position: number; node: PMNode }[] => {
  const paragraphs: { position: number; node: PMNode }[] = [];
  const { $from, $to } = state.selection;
  state.doc.nodesBetween($from.pos, $to.pos, (node, position) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    paragraphs.push({ position, node });
    return false;
  });
  return paragraphs;
};

/**
 * Rewrite the section record of every paragraph the selection touches.
 *
 * A record is rewritten once and the replacement applied to every paragraph
 * holding it, wherever in the document it sits. A split leaves two halves over
 * one object until the save leg decides which of them ends the section;
 * rewriting them apart would turn that one section into two, and leaving a
 * sharer outside the selection behind would do the same.
 *
 * A paragraph that holds no record is asked on its own, so two break-less
 * paragraphs in one selection author two sections rather than share one.
 */
const rewriteSelectedSectionRecords =
  (rewrite: SectionRecordRewrite): Command =>
  (state, dispatch) => {
    const selected = selectedParagraphs(state);
    if (selected.length === 0) {
      return false;
    }

    const replacements = new Map<SectionProperties, SectionProperties | null>();
    for (const { node } of selected) {
      const current = sectionPropertiesOf(node);
      if (current !== null && !replacements.has(current)) {
        replacements.set(current, rewrite(current));
      }
    }
    const selectedPositions = new Set(selected.map(({ position }) => position));

    const corrections: { position: number; properties: SectionProperties | null }[] = [];
    state.doc.descendants((node, position) => {
      if (node.type.name !== "paragraph") {
        return true;
      }
      const current = sectionPropertiesOf(node);
      if (current === null) {
        if (selectedPositions.has(position)) {
          const minted = rewrite(null);
          if (minted !== null) {
            corrections.push({ position, properties: minted });
          }
        }
        return false;
      }
      if (replacements.has(current)) {
        const replacement = replacements.get(current) ?? null;
        if (replacement !== current) {
          corrections.push({ position, properties: replacement });
        }
      }
      return false;
    });

    if (corrections.length === 0) {
      return false;
    }
    if (dispatch) {
      const transaction = state.tr;
      for (const { position, properties } of corrections) {
        const paragraph = transaction.doc.nodeAt(position);
        if (paragraph) {
          transaction.setNodeMarkup(
            position,
            undefined,
            withSectionProperties(paragraph.attrs, properties),
          );
        }
      }
      dispatch(transaction.scrollIntoView());
    }
    return true;
  };

/**
 * End the selection's sections with `breakType`.
 *
 * A paragraph that already ends a section keeps its page size, margins,
 * columns and header references and changes only `w:type` (§17.6.22), which is
 * the one field the type names; a paragraph that ends none mints a record
 * stating that type alone, and inherits the rest from the section that follows
 * it.
 */
export const setSectionBreakType = (breakType: SectionBreakType): Command =>
  rewriteSelectedSectionRecords((current) =>
    current === null ? mintSectionProperties(breakType) : { ...current, sectionStart: breakType },
  );

/**
 * Remove the section break the selection's paragraphs carry.
 *
 * The record goes, and with it the section boundary: the paragraphs it
 * governed become part of the section that follows, whose `w:sectPr` governs
 * them from then on. That is Word's own rule for deleting a section break, and
 * it is what §17.6.18 says a paragraph with no `w:sectPr` means — nothing has
 * to be merged into the following record, because the following record already
 * states the whole of the section it heads.
 */
export const removeSectionBreakAtSelection: Command = rewriteSelectedSectionRecords(() => null);

/**
 * Insert a section break at the current cursor position.
 *
 * In OOXML a section break is the `sectPr` carried by the *last* paragraph of a
 * section (§17.6.18). So we split the current paragraph at the cursor, give the
 * first half the one record this insertion mints (it becomes the section end),
 * and leave the cursor in the second half (the first paragraph of the new
 * section). Content after the cursor therefore flows into the new section —
 * onto a new page for `nextPage`, or in place for `continuous`.
 */
function insertSectionBreakAtCursor(breakType: InsertableSectionBreak): Command {
  return (state, dispatch) => {
    const { schema } = state;
    const paragraphType = schema.nodes["paragraph"];
    if (!paragraphType) return false;

    const { $from } = state.selection;
    // Section breaks only belong on top-level body paragraphs. Inside a table
    // cell or block SDT a `w:sectPr` is invalid OOXML, so refuse to act there
    // (the menu item becomes a no-op rather than corrupting the document).
    const isTopLevel = $from.parent.isTextblock
      ? state.doc.resolve($from.before()).depth === 0
      : $from.depth === 0;
    if (!isTopLevel) return false;

    if (dispatch) {
      const tr = state.tr;
      // One record per insertion: one break, one section, one object whose
      // identity a later split is read against.
      const sectionProperties = mintSectionProperties(breakType);
      let cursorPos: number;

      if ($from.parent.isTextblock) {
        // Position of the paragraph node the cursor sits in. Unaffected by the
        // split below, since the split happens at a later position.
        const paraPos = $from.before();
        tr.split($from.pos);
        const firstPara = tr.doc.nodeAt(paraPos);
        if (firstPara) {
          tr.setNodeMarkup(
            paraPos,
            undefined,
            withSectionProperties(firstPara.attrs, sectionProperties),
          );
        }
        // The split position maps to the start of the second paragraph's
        // content (the first paragraph of the new section) — put the cursor there.
        cursorPos = tr.mapping.map($from.pos);
      } else {
        // Not in a textblock — insert a section-ending empty paragraph here.
        // Place the cursor *inside* the new paragraph (pos + 1). Mapping
        // `$from.pos` would land on a block boundary, which is not a valid
        // `TextSelection` position and would throw.
        const pos = $from.pos;
        tr.insert(pos, paragraphType.create({ _sectionProperties: sectionProperties }));
        cursorPos = pos + 1;
      }

      tr.setSelection(TextSelection.create(tr.doc, cursorPos));
      dispatch(tr.scrollIntoView());
    }

    return true;
  };
}

/**
 * Insert a "next page" section break at the cursor — starts a new section on a
 * new page.
 */
export const insertSectionBreakNextPage: Command = insertSectionBreakAtCursor("nextPage");

/**
 * Insert a "continuous" section break at the cursor — starts a new section on
 * the same page.
 */
export const insertSectionBreakContinuous: Command = insertSectionBreakAtCursor("continuous");
