import { useCallback, useRef } from "react";

import type { EditorView } from "prosemirror-view";

import {
  createVisualLineState,
  handleVisualLineKeyDown,
} from "@stll/folio-core/prosemirror/utils/visualLineNavigation";

export type VisualLineNavigationOptions = {
  pagesContainerRef: React.RefObject<HTMLDivElement | null>;
};

/** Bind the framework-neutral visual-line navigation state to React. */
export function useVisualLineNavigation({ pagesContainerRef }: VisualLineNavigationOptions) {
  const stateRef = useRef(createVisualLineState());
  const handlePMKeyDown = useCallback(
    (view: EditorView, event: KeyboardEvent): boolean =>
      handleVisualLineKeyDown(stateRef.current, view, event, pagesContainerRef.current),
    [pagesContainerRef],
  );

  return { handlePMKeyDown };
}
