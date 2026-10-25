import { Selection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { splitsGraphemeCluster } from "../ai-edits/character-boundaries";
import { directionIsRtl } from "../prosemirror/paragraphDirection";

type ReadOnlySelectionView = Pick<EditorView, "dispatch" | "state">;
type ReadOnlySelectionKeyEvent = Pick<
  KeyboardEvent,
  "altKey" | "ctrlKey" | "key" | "metaKey" | "preventDefault" | "shiftKey"
>;

type TextPositionInDirectionOptions = {
  view: ReadOnlySelectionView;
  head: number;
  direction: -1 | 1;
};

const textPositionInDirection = ({
  view,
  head,
  direction,
}: TextPositionInDirectionOptions): number | null => {
  const $head = view.state.doc.resolve(head);
  const parent = $head.parent;
  let childOffset = 0;

  for (let childIndex = 0; childIndex < parent.childCount;) {
    const child = parent.child(childIndex);
    if (!child.isText) {
      childOffset += child.nodeSize;
      childIndex += 1;
      continue;
    }

    const runStart = childOffset;
    let runText = "";
    while (childIndex < parent.childCount && parent.child(childIndex).isText) {
      const textNode = parent.child(childIndex);
      runText += textNode.text ?? "";
      childOffset += textNode.nodeSize;
      childIndex += 1;
    }
    if ($head.parentOffset < runStart || $head.parentOffset > childOffset)
      continue;

    const localOffset = $head.parentOffset - runStart;
    let target = localOffset + direction;
    while (
      target > 0 &&
      target < runText.length &&
      splitsGraphemeCluster(runText, target)
    )
      target += direction;
    if (target >= 0 && target <= runText.length && target !== localOffset)
      return $head.start() + runStart + target;
    break;
  }

  const candidate = head + direction;
  if (candidate < 0 || candidate > view.state.doc.content.size) return null;
  return (
    Selection.findFrom(view.state.doc.resolve(candidate), direction, true)
      ?.head ?? null
  );
};

/**
 * Handle Ctrl/Cmd+A and unmodified horizontal Shift arrows in LTR paragraphs.
 * Word/document navigation, vertical movement, and RTL visual arrows stay
 * native because this helper does not model their platform-specific semantics.
 */
export const handleReadOnlySelectionKeyDown = (
  view: ReadOnlySelectionView,
  event: ReadOnlySelectionKeyEvent,
): boolean => {
  if (
    (event.metaKey || event.ctrlKey) &&
    !event.altKey &&
    event.key.toLowerCase() === "a"
  ) {
    event.preventDefault();
    const { doc } = view.state;
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(
          doc,
          Selection.atStart(doc).from,
          Selection.atEnd(doc).to,
        ),
      ),
    );
    return true;
  }

  if (
    !event.shiftKey ||
    event.altKey ||
    event.ctrlKey ||
    event.metaKey ||
    (event.key !== "ArrowLeft" && event.key !== "ArrowRight")
  )
    return false;

  const { doc, selection } = view.state;
  const direction = event.key === "ArrowLeft" ? -1 : 1;
  if (directionIsRtl(selection.$head.parent.attrs["direction"])) return false;
  const nextHead = textPositionInDirection({
    view,
    head: selection.head,
    direction,
  });
  if (nextHead === null) return false;

  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.between(
        doc.resolve(selection.anchor),
        doc.resolve(nextHead),
        direction,
      ),
    ),
  );
  event.preventDefault();
  return true;
};
