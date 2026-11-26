/**
 * A batch's comments run before its edits, and each comment reference is one
 * unit of its paragraph's offset space. An edit resolved against the
 * document before the batch reads its positions past the references the
 * batch's comments put at or before them, as the same edit reads them once
 * those comments stand.
 */
import {
  type CommentReferenceOffset,
  commentReferenceOffsets,
  type EditorIntent,
  idKey,
  type OpStory,
  paragraphsThroughTextBoxes,
  type TextPosition,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";

type ParagraphReferences = (
  document: Document,
  story: OpStory,
  blockId: string,
) => readonly CommentReferenceOffset[];

const paragraphReferences: ParagraphReferences = (document, story, blockId) => {
  const key = idKey(blockId);
  // A text box's paragraphs are named through the story drawing it.
  const paragraph = paragraphsThroughTextBoxes(document, story).find(
    ({ paragraph: candidate }) => idKey(candidate.paraId ?? "") === key,
  )?.paragraph;
  return paragraph === undefined ? [] : commentReferenceOffsets(paragraph);
};

type BatchReferenceShiftOptions = {
  /** The document before the batch. */
  before: Document;
  /** The document once the batch's comments stand, before its edits. */
  commented: Document;
};

/** Map a pre-batch position past the references the batch's comments added at or before it. */
export const batchCommentReferenceShift = ({
  before,
  commented,
}: BatchReferenceShiftOptions): ((position: TextPosition) => TextPosition) => {
  if (before === commented) return (position) => position;
  const added = new Map<string, readonly number[]>();
  const addedIn = (story: OpStory, blockId: string): readonly number[] => {
    const key = `${JSON.stringify(story)}\u0000${idKey(blockId)}`;
    const known = added.get(key);
    if (known !== undefined) return known;
    const kept = new Set(
      paragraphReferences(before, story, blockId).map(({ id }) => id),
    );
    // Each added reference's offset before the batch: its offset now, less
    // the added references before it.
    const offsets: number[] = [];
    for (const { id, offset } of paragraphReferences(commented, story, blockId))
      if (!kept.has(id)) offsets.push(offset - offsets.length);
    added.set(key, offsets);
    return offsets;
  };
  return (position) => {
    const offsets = addedIn(position.story, position.blockId);
    let shift = 0;
    for (const offset of offsets) if (offset <= position.offset) shift++;
    return shift === 0
      ? position
      : { ...position, offset: position.offset + shift };
  };
};

/** An intent with every text position it names mapped. */
export const mapIntentPositions = (
  intent: EditorIntent,
  map: (position: TextPosition) => TextPosition,
): EditorIntent => {
  switch (intent.type) {
    case "replaceText":
      return {
        ...intent,
        from: map(intent.from),
        to: map(intent.to),
        ...(intent.formattingSources === undefined
          ? {}
          : {
              formattingSources: intent.formattingSources.map((source) => ({
                ...source,
                at: map(source.at),
              })),
            }),
      };
    case "formatRun":
    case "insertAtom":
    case "replaceFragment":
    case "pasteBlocks":
    case "setHyperlink":
    case "removeHyperlink":
    case "insertHyperlink":
      return { ...intent, from: map(intent.from), to: map(intent.to) };
    case "moveFragment":
    case "moveBlocks":
      return {
        ...intent,
        from: map(intent.from),
        to: map(intent.to),
        target: map(intent.target),
      };
    case "splitParagraph":
      return {
        ...intent,
        at: map(intent.at),
        ...(intent.to === undefined ? {} : { to: map(intent.to) }),
      };
    case "formatParagraph":
    case "generateTOC":
      return { ...intent, at: map(intent.at) };
    case "setList":
      return {
        ...intent,
        items: intent.items.map((item) => ({ ...item, at: map(item.at) })),
      };
    case "table":
    case "tableRow":
    case "tableBlock":
    case "insertParagraphs":
    case "deleteParagraphs":
    case "joinParagraphs":
      return intent;
    default: {
      const unreachable: never = intent;
      return unreachable;
    }
  }
};
