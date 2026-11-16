/** Range endpoints whose deletion would unbalance a previously balanced range. */
import type { BlockContent, DocumentBody, Paragraph } from "../model/document";
import { compareGaps, type Gap, type LeafPlace, leafPlaces } from "./leaves";
import { rangeBoundaryDefinitions } from "./rangeBoundaries";

const BOUNDARY_SIDES = rangeBoundaryDefinitions();
/** Bookmarks may be unpaired; comment and move ranges must remain balanced. */
const BALANCED_FAMILIES: ReadonlySet<string> = new Set([
  "comment",
  "moveFrom",
  "moveTo",
]);

const hasBoundaryType = (type: string): type is keyof typeof BOUNDARY_SIDES =>
  Object.hasOwn(BOUNDARY_SIDES, type);

const balancedRangeOf = (node: { type: string; id?: unknown }) => {
  if (!hasBoundaryType(node.type)) return undefined;
  const { family, side } = BOUNDARY_SIDES[node.type];
  return BALANCED_FAMILIES.has(family)
    ? { key: `${family}:${String(node.id)}`, side }
    : undefined;
};

type Occurrence = { side: "start" | "end" } & (
  { type: "standing" } | { type: "removed"; paragraph: Paragraph; path: string }
);

const wellFormed = (occurrences: readonly Occurrence[]): boolean => {
  let open = 0;
  for (const { side } of occurrences) {
    open += side === "start" ? 1 : -1;
    if (open < 0) return false;
  }
  return open === 0;
};

type KeptDeletionBoundaryPathsOptions = {
  body: Pick<DocumentBody, "content">;
  removes: (leaf: LeafPlace, paragraph: Paragraph) => boolean;
};

/**
 * Source paths to retain, per paragraph. Paths identify occurrences even when
 * two positions hold the same immutable marker object. Complete source-ordered
 * groups are removed independently; malformed ranges are not repaired.
 */
export const keptDeletionBoundaryPaths = ({
  body,
  removes,
}: KeptDeletionBoundaryPathsOptions): Map<Paragraph, Set<string>> => {
  const ranges = new Map<string, Occurrence[]>();
  const record = (
    range: NonNullable<ReturnType<typeof balancedRangeOf>>,
    occurrence: Occurrence,
  ): void => {
    const occurrences = ranges.get(range.key) ?? [];
    occurrences.push(occurrence);
    ranges.set(range.key, occurrences);
  };
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      const range = balancedRangeOf(block);
      if (range !== undefined) {
        record(range, { type: "standing", side: range.side });
        continue;
      }
      switch (block.type) {
        case "paragraph":
          for (const leaf of leafPlaces(block.content)) {
            const boundary = balancedRangeOf(leaf.node);
            if (boundary === undefined) continue;
            record(
              boundary,
              removes(leaf, block)
                ? {
                    type: "removed",
                    side: boundary.side,
                    paragraph: block,
                    path: leaf.path.join("."),
                  }
                : { type: "standing", side: boundary.side },
            );
          }
          break;
        case "table":
          for (const row of block.rows)
            for (const cell of row.cells) visit(cell.content);
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content);
          break;
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          break;
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    }
  };
  visit(body.content);
  const kept = new Map<Paragraph, Set<string>>();
  for (const occurrences of ranges.values()) {
    if (!wellFormed(occurrences)) continue;
    // Separate source-ordered pairs remain separate ranges even when they
    // share an ID. Nested duplicate endpoints stay one ambiguous group.
    let group: Occurrence[] = [];
    let depth = 0;
    for (const occurrence of occurrences) {
      group.push(occurrence);
      depth += occurrence.side === "start" ? 1 : -1;
      if (depth !== 0) continue;
      if (!wellFormed(group.filter(({ type }) => type === "standing")))
        for (const removed of group) {
          if (removed.type !== "removed") continue;
          const paths = kept.get(removed.paragraph) ?? new Set<string>();
          paths.add(removed.path);
          kept.set(removed.paragraph, paths);
        }
      group = [];
    }
  }
  return kept;
};

type KeptDeletionBoundariesOptions = {
  body: Pick<DocumentBody, "content">;
  paragraph: Paragraph;
  from: Gap;
  to: Gap;
};

export const keptDeletionBoundaries = ({
  body,
  paragraph,
  from,
  to,
}: KeptDeletionBoundariesOptions): LeafPlace[] => {
  const leaves = leafPlaces(paragraph.content);
  const covered = leaves.filter(
    ({ node, before, after }) =>
      balancedRangeOf(node) !== undefined &&
      compareGaps(from, before) <= 0 &&
      compareGaps(after, to) <= 0,
  );
  if (covered.length === 0) return [];
  const paths = keptDeletionBoundaryPaths({
    body,
    removes: (leaf, owner) =>
      owner === paragraph &&
      compareGaps(from, leaf.before) <= 0 &&
      compareGaps(leaf.after, to) <= 0,
  }).get(paragraph);
  return paths === undefined
    ? []
    : covered.filter(({ path }) => paths.has(path.join(".")));
};
