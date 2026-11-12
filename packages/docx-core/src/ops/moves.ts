/** Tracked block insertions and deletions recorded as the two halves of a move. */
import type { ParagraphContent } from "../model/document";

type MoveHalf = "from" | "to";

const RENAMED = {
  from: { record: "deletion", move: "moveFrom" },
  to: { record: "insertion", move: "moveTo" },
} as const satisfies Record<MoveHalf, { record: string; move: string }>;

/**
 * `content` with every record the operation's own revision `revisionId`
 * created renamed to the move half it is: a deletion to `w:moveFrom`, an
 * insertion to `w:moveTo`. Records of other revisions are left as they are.
 */
export const asMoveHalf = (
  content: readonly ParagraphContent[],
  revisionId: number,
  half: MoveHalf,
): ParagraphContent[] =>
  content.map((item) => movedItem(item, revisionId, half));

const movedItem = <T>(item: T, revisionId: number, half: MoveHalf): T => {
  if (typeof item !== "object" || item === null) return item;
  if (Array.isArray(item))
    // The mapped array holds the same element types as `item`.
    return item.map((child: unknown) =>
      movedItem(child, revisionId, half),
    ) as T;
  const record = item as Record<string, unknown>;
  let next: Record<string, unknown> = record;
  const content = record["content"];
  if (Array.isArray(content))
    next = { ...next, content: movedItem(content, revisionId, half) };
  const children = record["children"];
  if (Array.isArray(children))
    next = { ...next, children: movedItem(children, revisionId, half) };
  const info = record["info"] as { id?: unknown } | undefined;
  if (record["type"] === RENAMED[half].record && info?.id === revisionId)
    next = { ...next, type: RENAMED[half].move };
  // The renamed record has the fields the move record has.
  return (next === record ? item : next) as T;
};
