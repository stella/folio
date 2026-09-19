/**
 * Count body references whose relationship id names nothing in the package.
 *
 * A dangling id is a fact about the source, not a reason to drop content: Word
 * opens such a package, so folio must too. What must never happen is the id
 * quietly resolving to some other part, and the way that defect hides is that
 * a miss looks exactly like an id the author never wrote. Counting the misses
 * turns the two apart and puts the difference in `document.warnings`.
 *
 * Scoped to the main story and the document relationships. A header, footer or
 * notes part carries its own `.rels`, which the parsed model does not keep, so
 * a reference inside one cannot be checked from the model.
 */

import type { BlockContent, RelationshipMap } from "../types/document";
import { resolveRelationshipId } from "./relsParser";

export type DanglingRelationshipReferences = {
  /** Drawings whose `r:embed` names no relationship. */
  drawings: number;
  /** Hyperlinks whose `r:id` names no relationship. */
  hyperlinks: number;
};

type CountDanglingRelationshipReferencesOptions = {
  content: readonly BlockContent[];
  relationships: RelationshipMap | undefined;
};

export const countDanglingRelationshipReferences = ({
  content,
  relationships,
}: CountDanglingRelationshipReferencesOptions): DanglingRelationshipReferences => {
  const counts: DanglingRelationshipReferences = { drawings: 0, hyperlinks: 0 };
  const isDangling = (rId: string | undefined): boolean =>
    resolveRelationshipId(relationships, rId).status === "dangling";

  const visitBlocks = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            visitBlocks(cell.content);
          }
        }
        continue;
      }
      if (block.type !== "paragraph") {
        continue;
      }
      for (const item of block.content) {
        if (item.type === "hyperlink") {
          if (isDangling(item.rId)) {
            counts.hyperlinks += 1;
          }
          continue;
        }
        if (item.type !== "run") {
          continue;
        }
        for (const child of item.content) {
          if (child.type === "drawing" && isDangling(child.image.rId)) {
            counts.drawings += 1;
          }
        }
      }
    }
  };

  visitBlocks(content);
  return counts;
};
