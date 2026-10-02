import { structurallyEqual } from "../equality";
import { OP_STORIES, type OpStory } from "../types";

const DECLARED_STORIES = Object.values(OP_STORIES) satisfies readonly OpStory[];

/** Compare story identities without depending on their representation. */
export const sameStory = (left: OpStory, right: OpStory): boolean => structurallyEqual(left, right);

/** The wire decoder accepts only stories declared by the operation contract. */
export const isOpStory = (value: unknown): value is OpStory =>
  DECLARED_STORIES.some((story) => structurallyEqual(story, value));
