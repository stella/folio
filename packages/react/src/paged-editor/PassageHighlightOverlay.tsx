/**
 * Passage Highlight Overlay
 *
 * Paints a single persistent, translucent highlight over a text passage a
 * consumer opened the document at (a citation chip's quoted passage, a
 * find-in-document hit, an agent tool). The highlighted range is resolved in
 * the hidden ProseMirror editor and projected onto container-space rectangles
 * (via {@link projectRangesToRects} in the host), so this component stays dumb:
 * it just stacks one absolutely-positioned span per line rectangle.
 *
 * Pointer events stay disabled so the caret, selection, and the editor's own
 * handlers keep working through the overlay. The highlight is view state only:
 * the host clears it on any doc-changing transaction and via the explicit clear
 * method, so this overlay never has to reason about the document itself.
 */

import type { CSSProperties } from "react";

import type { SelectionRect } from "@stll/folio-core/layout-bridge/engine/selectionRects";

export type PassageHighlightOverlayProps = {
  rects: SelectionRect[];
};

const overlayStyles: CSSProperties = {
  position: "absolute",
  top: 0,
  left: "50%",
  width: "100vw",
  height: "100%",
  transform: "translateX(-50%)",
  pointerEvents: "none",
  zIndex: 1,
};

export const PassageHighlightOverlay = ({ rects }: PassageHighlightOverlayProps) => {
  if (rects.length === 0) {
    return null;
  }

  return (
    <div style={overlayStyles} data-folio-passage-highlight-overlay="">
      {rects.map((rect, idx) => (
        <span
          key={`${idx}:${rect.pageIndex}:${rect.x}:${rect.y}`}
          className="folio-passage-highlight"
          style={{
            position: "absolute",
            left: rect.x,
            top: rect.y,
            width: rect.width,
            height: rect.height,
          }}
        />
      ))}
    </div>
  );
};
