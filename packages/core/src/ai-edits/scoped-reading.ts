import { isFolioAIContentBlock } from "./snapshot";
import type {
  FolioAIEditSnapshot,
  FolioDocumentOutline,
  FolioDocumentOutlineEntry,
  FolioDocumentSectionHandle,
  FolioDocumentSectionReadResult,
} from "./types";

const toSectionHandle = (
  snapshot: FolioAIEditSnapshot,
  headingBlockId: string,
  headingLevel: number,
): FolioDocumentSectionHandle | null => {
  const headingTextHash = snapshot.anchors[headingBlockId]?.textHash;
  if (headingTextHash === undefined) {
    return null;
  }
  return {
    type: "headingSection",
    story: "main",
    headingBlockId,
    headingTextHash,
    headingLevel,
  };
};

/** Build a flat, ordered outline with stable handles and explicit parents. */
export const getFolioDocumentOutline = (snapshot: FolioAIEditSnapshot): FolioDocumentOutline => {
  const sections: FolioDocumentOutlineEntry[] = [];
  const parentStack: FolioDocumentOutlineEntry[] = [];

  for (const block of snapshot.blocks) {
    // A heading-styled blank paragraph is not a section: it has no text to
    // name it, and every one of them would hash alike and collide as handles.
    if (block.headingLevel === undefined || !isFolioAIContentBlock(block)) {
      continue;
    }
    const handle = toSectionHandle(snapshot, block.id, block.headingLevel);
    if (handle === null) {
      continue;
    }

    while ((parentStack.at(-1)?.level ?? 0) >= block.headingLevel) {
      parentStack.pop();
    }
    const parentHandle = parentStack.at(-1)?.handle;
    const entry: FolioDocumentOutlineEntry = {
      handle,
      headingBlockId: block.id,
      text: block.text,
      level: block.headingLevel,
      ...(parentHandle !== undefined && { parentHandle }),
    };
    sections.push(entry);
    parentStack.push(entry);
  }

  return { sections };
};

/** Resolve one heading section against a fresh snapshot without guessing. */
export const readFolioDocumentSection = (
  snapshot: FolioAIEditSnapshot,
  handle: FolioDocumentSectionHandle,
): FolioDocumentSectionReadResult => {
  const outline = getFolioDocumentOutline(snapshot);
  const heading = outline.sections.find(
    ({ headingBlockId }) => headingBlockId === handle.headingBlockId,
  );
  if (heading === undefined) {
    return { status: "missing" };
  }
  if (
    heading.handle.headingTextHash !== handle.headingTextHash ||
    heading.level !== handle.headingLevel
  ) {
    return { status: "stale" };
  }

  const startIndex = snapshot.blocks.findIndex(({ id }) => id === handle.headingBlockId);
  if (startIndex === -1) {
    return { status: "missing" };
  }

  // The same test the outline applies, or the two disagree about what a
  // section is: a heading-styled BLANK paragraph is not one, and treating it
  // as a boundary here ended the section early and dropped everything the
  // outline still counts as inside it.
  let endIndex = snapshot.blocks.length;
  for (let index = startIndex + 1; index < snapshot.blocks.length; index++) {
    const block = snapshot.blocks.at(index);
    if (
      block !== undefined &&
      block.headingLevel !== undefined &&
      block.headingLevel <= heading.level &&
      isFolioAIContentBlock(block)
    ) {
      endIndex = index;
      break;
    }
  }

  return {
    status: "found",
    section: {
      handle,
      heading,
      // A section is what a reader would read. The blank paragraphs between
      // its blocks are part of the document's shape, not part of the section.
      blocks: snapshot.blocks.slice(startIndex, endIndex).filter(isFolioAIContentBlock),
    },
  };
};
