import { panic } from "better-result";
import type { ResolvedPos } from "prosemirror-model";

type ParagraphMarkContainerEndOptions = {
  at: ResolvedPos;
  /** The whole-story walk has already removed every following local child. */
  ending: "paragraph" | "parent";
};

/**
 * The logical file container the paragraph at `at` ends. Content controls and
 * custom XML are transparent block-list steps, as in the kernel's
 * `endsItsContainer`; a cell terminator is retained even when an earlier
 * paragraph could end that cell. Editor-only text boxes belong to their
 * anchor paragraph's run rather than the surrounding block list.
 */
export const paragraphMarkContainerEnd = ({
  at,
  ending,
}: ParagraphMarkContainerEndOptions): "interior" | "story-end" | "cell-end" => {
  for (let depth = at.depth; depth >= 0; depth -= 1) {
    const parent = at.node(depth);
    let last = parent.childCount - 1;
    while (parent.maybeChild(last)?.type.name === "textBox") last -= 1;
    if (
      (depth !== at.depth || ending === "paragraph") &&
      at.index(depth) !== last
    )
      return "interior";
    const role = parent.type.spec["tableRole"];
    if (role === "cell" || role === "header_cell") return "cell-end";
    if (
      parent.type.name !== "blockSdt" &&
      parent.type.name !== "blockCustomXml"
    )
      return "story-end";
  }
  return panic("A paragraph is enclosed by a logical story or cell.");
};
