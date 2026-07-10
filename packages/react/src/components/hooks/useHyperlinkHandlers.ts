import { useCallback, useState } from "react";

import type { EditorView } from "prosemirror-view";

import {
  editHyperlinkAtCursor,
  removeHyperlinkAtCursor,
} from "@stll/folio-core/prosemirror/commands/hyperlink";
import { sanitizeExternalUrl } from "@stll/folio-core/utils/urlSecurity";
import type { HyperlinkPopupData } from "../ui/HyperlinkPopup";
import { toast } from "../toast";

// ============================================================================
// TYPES
// ============================================================================

export type UseHyperlinkHandlersDeps = {
  /** Returns the currently active ProseMirror editor view */
  getActiveEditorView: () => EditorView | null | undefined;
  /** Focuses the currently active editor */
  focusActiveEditor: () => void;
};

export type UseHyperlinkHandlersReturn = {
  /** Popup state (Google Docs-style floating popup on link click) */
  hyperlinkPopupData: HyperlinkPopupData | null;
  /** Update popup state directly */
  setHyperlinkPopupData: (data: HyperlinkPopupData | null) => void;
  /** Handle hyperlink click to show popup */
  handleHyperlinkClick: (data: HyperlinkPopupData) => void;
  /** Navigate to the link URL in a new tab */
  handleHyperlinkPopupNavigate: (href: string) => void;
  /** Copy the link URL to clipboard */
  handleHyperlinkPopupCopy: (href: string) => void;
  /** Edit the hyperlink text and URL inline */
  handleHyperlinkPopupEdit: (displayText: string, href: string) => void;
  /** Remove the hyperlink mark from the popup */
  handleHyperlinkPopupRemove: () => void;
  /** Close the hyperlink popup */
  handleHyperlinkPopupClose: () => void;
};

// ============================================================================
// HOOK
// ============================================================================

export const useHyperlinkHandlers = ({
  getActiveEditorView,
  focusActiveEditor,
}: UseHyperlinkHandlersDeps): UseHyperlinkHandlersReturn => {
  const [hyperlinkPopupData, setHyperlinkPopupData] = useState<HyperlinkPopupData | null>(null);

  // Handle hyperlink click — show popup
  const handleHyperlinkClick = useCallback(
    (data: HyperlinkPopupData) => setHyperlinkPopupData(data),
    [],
  );

  const handleHyperlinkPopupNavigate = useCallback((href: string) => {
    const safeHref = sanitizeExternalUrl(href);
    if (safeHref) {
      window.open(safeHref, "_blank", "noopener,noreferrer");
    }
  }, []);

  const handleHyperlinkPopupCopy = useCallback((href: string) => {
    void navigator.clipboard.writeText(href);
  }, []);

  const handleHyperlinkPopupEdit = useCallback(
    (displayText: string, href: string) => {
      const view = getActiveEditorView();
      if (!view) {
        return;
      }

      if (!editHyperlinkAtCursor(view, { displayText, href })) {
        return;
      }

      setHyperlinkPopupData(null);
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor],
  );

  const handleHyperlinkPopupRemove = useCallback(() => {
    const view = getActiveEditorView();
    if (!view) {
      return;
    }

    if (!removeHyperlinkAtCursor(view, { popupHref: hyperlinkPopupData?.href })) {
      return;
    }

    setHyperlinkPopupData(null);
    focusActiveEditor();
    toast("Link removed");
  }, [getActiveEditorView, focusActiveEditor, hyperlinkPopupData]);

  const handleHyperlinkPopupClose = useCallback(() => {
    setHyperlinkPopupData(null);
  }, []);

  return {
    hyperlinkPopupData,
    setHyperlinkPopupData,
    handleHyperlinkClick,
    handleHyperlinkPopupNavigate,
    handleHyperlinkPopupCopy,
    handleHyperlinkPopupEdit,
    handleHyperlinkPopupRemove,
    handleHyperlinkPopupClose,
  };
};
