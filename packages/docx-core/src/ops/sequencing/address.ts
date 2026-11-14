import { isRelationshipId } from "../../model/relationshipId";
import { structurallyEqual } from "../equality";
import { OP_STORIES, type OpStory } from "../types";

/** Compare story identities without depending on their representation. */
export const sameStory = (left: OpStory, right: OpStory): boolean =>
  structurallyEqual(left, right);

/** Whether `value` holds exactly the fields `keys` names. */
const hasExactly = (value: object, keys: readonly string[]): boolean =>
  Object.keys(value).length === keys.length &&
  keys.every((key) => Object.hasOwn(value, key));

/**
 * The wire decoder accepts every story the operation contract names: the
 * main story, a header or footer part by its relationship id, and a note by
 * its id. Nothing else, and no extra field, passes.
 */
export const isOpStory = (value: unknown): value is OpStory => {
  if (value === OP_STORIES.MAIN) return true;
  if (typeof value !== "object" || value === null || !("kind" in value))
    return false;
  const { kind } = value;
  if ((kind === "header" || kind === "footer") && "rId" in value)
    return (
      hasExactly(value, ["kind", "rId"]) &&
      typeof value.rId === "string" &&
      isRelationshipId(value.rId)
    );
  if ((kind === "footnote" || kind === "endnote") && "id" in value)
    return (
      hasExactly(value, ["kind", "id"]) &&
      typeof value.id === "number" &&
      Number.isSafeInteger(value.id) &&
      value.id >= 0
    );
  return false;
};
