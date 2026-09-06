import type { FolioDocumentStoryHandle } from "./ai-edits/headless";

export type FolioDocumentStoryPair = {
  baseStory: FolioDocumentStoryHandle | null;
  revisedStory: FolioDocumentStoryHandle | null;
};

const documentStoryKey = (story: FolioDocumentStoryHandle): string => {
  if (story.type === "main") {
    return story.type;
  }
  if (story.type === "header" || story.type === "footer") {
    return `${story.type}:${story.relationshipId}`;
  }
  return `${story.type}:${String(story.noteId)}`;
};

/**
 * Pair two packages' stories: by identity first, then, for headers and
 * footers, by kind and document order.
 *
 * A relationship id names a part inside ONE package. Two revisions of one file
 * carry it forward, so matching on it pairs the right header exactly and is
 * tried first. Two independently authored files do not: `rId6` in one has
 * nothing to do with `rId6` in the other, and matching on it alone left every
 * header and footer of such a pair reported as present on one side only —
 * a comparison of two contracts covered their bodies and nothing else.
 *
 * Order is what is left to pair on. It is the order the package lists its
 * parts in, so the nth header of one document answers to the nth header of the
 * other, and a document with more of them than the other leaves its surplus
 * unpaired rather than pairing a first-page header against a default one.
 */
export const pairFolioDocumentStories = (
  baseStories: readonly FolioDocumentStoryHandle[],
  revisedStories: readonly FolioDocumentStoryHandle[],
): FolioDocumentStoryPair[] => {
  const revisedByKey = new Map(revisedStories.map((story) => [documentStoryKey(story), story]));
  const pairedKeys = new Set<string>();
  const pairedByIdentity = new Map<FolioDocumentStoryHandle, FolioDocumentStoryHandle>();
  for (const baseStory of baseStories) {
    const key = documentStoryKey(baseStory);
    const revisedStory = revisedByKey.get(key);
    if (revisedStory) {
      pairedKeys.add(key);
      pairedByIdentity.set(baseStory, revisedStory);
    }
  }

  const unpairedRevisedOfKind = (kind: "header" | "footer"): FolioDocumentStoryHandle[] =>
    revisedStories.filter(
      (story) => story.type === kind && !pairedKeys.has(documentStoryKey(story)),
    );
  const remaining = {
    header: unpairedRevisedOfKind("header"),
    footer: unpairedRevisedOfKind("footer"),
  };

  const pairs: FolioDocumentStoryPair[] = [];
  for (const baseStory of baseStories) {
    const identityMatch = pairedByIdentity.get(baseStory);
    if (identityMatch) {
      pairs.push({ baseStory, revisedStory: identityMatch });
      continue;
    }
    if (baseStory.type !== "header" && baseStory.type !== "footer") {
      pairs.push({ baseStory, revisedStory: null });
      continue;
    }
    const byOrder = remaining[baseStory.type].shift() ?? null;
    if (byOrder) {
      pairedKeys.add(documentStoryKey(byOrder));
    }
    pairs.push({ baseStory, revisedStory: byOrder });
  }

  for (const revisedStory of revisedStories) {
    if (!pairedKeys.has(documentStoryKey(revisedStory))) {
      pairs.push({ baseStory: null, revisedStory });
    }
  }
  return pairs;
};
