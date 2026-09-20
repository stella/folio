/**
 * Retarget pasted heading paragraphs at the open document's own heading style.
 *
 * The schema's `h1`–`h6` paste rule runs inside `parseDOM`, which sees no
 * document, so it can only write the English built-in id `Heading1`…`Heading6`.
 * In a document written by a localized Word the heading style is `Nadpis1` or
 * `berschrift2`, and in one folio generated from a style set the deeper levels
 * may not exist at all, so that id resolves to nothing and the pasted heading
 * loses the document's heading formatting.
 *
 * The paragraph still classifies as a heading either way — `parseDOM` also sets
 * `outlineLevel`, which is what `docx/builtInStyles.ts` reads first. This is
 * about which style the text picks up.
 *
 * Deliberately narrow: a paragraph is retargeted only when the document defines
 * no style under the id it carries. Pasting between two folio documents, or
 * within one, never reaches it. Creating a style stays an explicit command —
 * minting one here would edit `styles.xml` as a side effect of a clipboard
 * action — so a level the document does not define falls back to the deepest
 * heading it does. An `<h5>` pasted into a set that stops at `heading 4` lands
 * on that one rather than keeping an id nothing resolves.
 */

import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

import { headingLevelOf } from "@stll/docx-core/model";

import type { BuiltInStyleIndex } from "../../../docx/builtInStyles";
import { readOutlineLevelAttr } from "../../outlineLevelAttr";
import { getDocumentBuiltInStyles, getDocumentStyleResolver } from "../../plugins/documentStyles";

/**
 * The document's style for this heading level, else its deepest shallower one.
 * Never deeper: promoting an `<h5>` to `heading 1` would restructure the
 * document's outline, while demoting it to the deepest level it has preserves
 * the ordering of everything around it.
 */
const nearestHeadingStyleId = (
  outlineLevel: number,
  styles: BuiltInStyleIndex,
): string | undefined => {
  for (let level = outlineLevel; level >= 0; level -= 1) {
    const styleId = styles.styleIdForHeadingLevel(level);
    if (styleId !== undefined) {
      return styleId;
    }
  }
  return undefined;
};

const retargetedStyleId = (node: PMNode, view: EditorView): string | undefined => {
  const styleId: unknown = node.attrs["styleId"];
  if (typeof styleId !== "string") {
    return undefined;
  }
  const headingLevel = headingLevelOf(readOutlineLevelAttr(node.attrs["outlineLevel"]));
  if (headingLevel === undefined) {
    return undefined;
  }
  const resolver = getDocumentStyleResolver(view.state);
  if (resolver === null || resolver.getStyle(styleId) !== undefined) {
    return undefined;
  }
  const target = nearestHeadingStyleId(headingLevel, getDocumentBuiltInStyles(view.state));
  return target === styleId ? undefined : target;
};

const retargetFragment = (fragment: Fragment, view: EditorView): Fragment => {
  const children: PMNode[] = [];
  let changed = false;
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Fragment.forEach
  fragment.forEach((node) => {
    const content = retargetFragment(node.content, view);
    const styleId = node.type.name === "paragraph" ? retargetedStyleId(node, view) : undefined;
    if (styleId === undefined && content === node.content) {
      children.push(node);
      return;
    }
    changed = true;
    const retargeted =
      styleId === undefined
        ? node
        : node.type.create({ ...node.attrs, styleId }, content, node.marks);
    children.push(styleId === undefined ? node.copy(content) : retargeted);
  });
  return changed ? Fragment.fromArray(children) : fragment;
};

/** A `transformPasted` hook; returns the slice unchanged when nothing matches. */
export const retargetPastedHeadingStyles = (slice: Slice, view: EditorView): Slice => {
  const content = retargetFragment(slice.content, view);
  return content === slice.content ? slice : new Slice(content, slice.openStart, slice.openEnd);
};
