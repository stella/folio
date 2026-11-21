/**
 * Whose pending record an edit takes back. Deleting over one's own pending
 * insertion removes it outright, as no one else needs it back; a record of
 * another author, or of another suggestion (an ordinary tracked change
 * included), stays and the deletion is recorded beside it, so rejecting the
 * deletion gives it back.
 */
import type { RevisionStamp } from "./types";

/** The author of an edit and, for a pending suggestion, the suggestion it belongs to. */
export type RevisionOwner = Pick<RevisionStamp, "author" | "suggestionId">;

type RecordInfo = {
  author?: string | undefined;
  suggestionId?: string | undefined;
};

/** Whether `owner` made the record: the same author, and the same suggestion or none on both. */
export const isOwnRecord = (info: RecordInfo, owner: RevisionOwner): boolean =>
  info.author === owner.author && info.suggestionId === owner.suggestionId;

/** Every tracked record in `value`, at any depth, is `owner`'s. */
export const recordedOnlyBy = (
  value: unknown,
  owner: RevisionOwner,
): boolean => {
  if (Array.isArray(value))
    return value.every((item) => recordedOnlyBy(item, owner));
  if (typeof value !== "object" || value === null) return true;
  const info: unknown = Reflect.get(value, "info");
  if (typeof info === "object" && info !== null) {
    const author: unknown = Reflect.get(info, "author");
    const suggestionId: unknown = Reflect.get(info, "suggestionId");
    if (
      author !== undefined &&
      (author !== owner.author || suggestionId !== owner.suggestionId)
    )
      return false;
  }
  return Object.values(value).every((item) => recordedOnlyBy(item, owner));
};
