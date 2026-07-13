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
  };
};

/** Build a flat, ordered outline with stable handles and explicit parents. */
export const getFolioDocumentOutline = (snapshot: FolioAIEditSnapshot): FolioDocumentOutline => {
  const sections: FolioDocumentOutlineEntry[] = [];
  const parentStack: FolioDocumentOutlineEntry[] = [];

  for (const block of snapshot.blocks) {
    if (block.headingLevel === undefined) {
      continue;
    }
    const handle = toSectionHandle(snapshot, block.id);
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
  if (heading.handle.headingTextHash !== handle.headingTextHash) {
    return { status: "stale" };
  }

  const startIndex = snapshot.blocks.findIndex(({ id }) => id === handle.headingBlockId);
  if (startIndex === -1) {
    return { status: "missing" };
  }

  let endIndex = snapshot.blocks.length;
  for (let index = startIndex + 1; index < snapshot.blocks.length; index++) {
    const block = snapshot.blocks.at(index);
    if (block?.headingLevel !== undefined && block.headingLevel <= heading.level) {
      endIndex = index;
      break;
    }
  }

  return {
    status: "found",
    section: {
      handle,
      heading,
      blocks: snapshot.blocks.slice(startIndex, endIndex),
    },
  };
};
