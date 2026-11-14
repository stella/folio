import { structurallyEqual } from "../equality";
import { OP_STORIES, type OpStory } from "../types";
import {
  literal,
  natural,
  nonemptyString,
  object,
  type Validator,
} from "./wireRecords";

/** Compare story identities without depending on their representation. */
export const sameStory = (left: OpStory, right: OpStory): boolean =>
  structurallyEqual(left, right);

/** A header or footer part's story, by its relationship id. */
export const headerFooterStory: Validator = object({
  kind: literal("header", "footer"),
  rId: nonemptyString,
});
/** A footnote's or an endnote's story, by its id. */
export const noteStory: Validator = object({
  kind: literal("footnote", "endnote"),
  id: natural,
});

/**
 * The wire decoder accepts every story the operation contract names: the
 * main story, a header or footer part by its relationship id, and a note by
 * its id. Nothing else, and no extra field, passes.
 */
export const isOpStory = (value: unknown): value is OpStory =>
  value === OP_STORIES.MAIN || headerFooterStory(value) || noteStory(value);
