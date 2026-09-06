/**
 * State threaded through one `buildDisplayList` call.
 *
 * Everything mutable the builder needs lives here so the builder itself stays a
 * pure function of its arguments: the font table, the image table, the
 * unsupported log and the author-colour assignment are all born and die inside
 * a single call. `utils/authorColors.ts`'s `getAuthorColorIdx` deliberately is
 * not used: it keeps a module-level map, so the colour an author gets depends
 * on which documents the process opened earlier, and two builds of the same
 * layout would not compare equal.
 */

import { AUTHOR_COLORS } from "../../utils/authorColors";
import type { DisplayColor, DisplayLink, DisplayLinkTarget } from "../types";
import { parseDisplayColor, SUGGESTION_COLOR } from "./colors";
import type { FontTable } from "./fontTable";
import type { ImageTable } from "./imagePrimitives";
import type { UnsupportedCollector } from "./unsupported";

/** Word's per-author redline palette, assigned in first-seen order per build. */
export class AuthorColorTable {
  private readonly indexByAuthor = new Map<string, number>();

  colorFor(author: string | undefined): DisplayColor {
    const key = author ?? "";
    let index = this.indexByAuthor.get(key);
    if (index === undefined) {
      index = this.indexByAuthor.size % AUTHOR_COLORS.length;
      this.indexByAuthor.set(key, index);
    }
    // SAFETY: index is `size % AUTHOR_COLORS.length`, always in bounds.
    return parseDisplayColor(AUTHOR_COLORS[index]!) ?? SUGGESTION_COLOR;
  }
}

/** Colour a tracked change paints in: the author's hue, or the proposal hue. */
export const trackedChangeColor = (
  authorColors: AuthorColorTable,
  changeAuthor: string | undefined,
  isSuggestion: boolean | undefined,
): DisplayColor => (isSuggestion ? SUGGESTION_COLOR : authorColors.colorFor(changeAuthor));

/**
 * Which OOXML story the primitives being built belong to.
 *
 * Only the body's runs carry positions in the editable model the caller holds.
 * A header, a footer and a footnote are each their own story with their own
 * document positions, so a range taken from one of them would address
 * unrelated body content; the painter drops the same anchors for the same
 * reason (`stripFootnotePmAnchors`).
 */
export type DisplayStory = "body" | "header" | "footer" | "footnote";

export type BuildContext = {
  readonly fonts: FontTable;
  readonly images: ImageTable;
  readonly unsupported: UnsupportedCollector;
  readonly authorColors: AuthorColorTable;
  /** Link sink for the page being built. */
  readonly links: DisplayLink[];
  /**
   * Bookmark name → where it lands, collected in a first pass over the layout.
   * A `#name` hyperlink can only become a `page` target once every page is
   * known, which is why this is resolved before any page is painted.
   */
  readonly bookmarkTargets: ReadonlyMap<string, DisplayLinkTarget>;
  readonly story: DisplayStory;
  readonly pageIndex: number;
  /** Authored page number, for `PAGE` fields. */
  readonly pageNumber: number;
  readonly totalPages: number;
};
