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

export const pairFolioDocumentStories = (
  baseStories: readonly FolioDocumentStoryHandle[],
  revisedStories: readonly FolioDocumentStoryHandle[],
): FolioDocumentStoryPair[] => {
  const revisedByKey = new Map(revisedStories.map((story) => [documentStoryKey(story), story]));
  const pairedKeys = new Set<string>();
  const pairs: FolioDocumentStoryPair[] = [];
  for (const baseStory of baseStories) {
    const key = documentStoryKey(baseStory);
    const revisedStory = revisedByKey.get(key) ?? null;
    pairs.push({ baseStory, revisedStory });
    if (revisedStory) {
      pairedKeys.add(key);
    }
  }
  for (const revisedStory of revisedStories) {
    if (!pairedKeys.has(documentStoryKey(revisedStory))) {
      pairs.push({ baseStory: null, revisedStory });
    }
  }
  return pairs;
};
