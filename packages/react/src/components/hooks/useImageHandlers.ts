import { useCallback, useState } from "react";

import type { EditorView } from "prosemirror-view";

import {
  applyImagePosition,
  applyImageProperties,
  applyImageTransform,
  applyImageWrapType,
  insertImageFromFile,
} from "@stll/folio-core/prosemirror/commands/image";
import type { ImageTransformAction } from "@stll/folio-core/prosemirror/commands/image";

import type { ImagePositionData } from "../dialogs/ImagePositionDialog";
import type { ImagePropertiesData } from "../dialogs/ImagePropertiesDialog";

// ============================================================================
// TYPES
// ============================================================================

export type ImageContext = {
  pos: number;
  wrapType: string;
  displayMode: string;
  cssFloat: string | null;
  transform: string | null;
  alt: string | null;
  borderWidth: number | null;
  borderColor: string | null;
  borderStyle: string | null;
};

export type UseImageHandlersDeps = {
  /** Returns the currently active ProseMirror editor view */
  getActiveEditorView: () => EditorView | null | undefined;
  /** Focuses the currently active editor */
  focusActiveEditor: () => void;
  /** Image context when cursor is on an image node */
  pmImageContext: ImageContext | null;
};

export type UseImageHandlersReturn = {
  /** Whether the image position dialog is open */
  imagePositionOpen: boolean;
  /** Set image position dialog open state */
  setImagePositionOpen: (open: boolean) => void;
  /** Whether the image properties dialog is open */
  imagePropsOpen: boolean;
  /** Set image properties dialog open state */
  setImagePropsOpen: (open: boolean) => void;
  /** Handle file input change for image insert */
  handleImageFileChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  /** Handle image wrap type change */
  handleImageWrapType: (wrapType: string) => void;
  /** Handle image transform (rotate/flip) */
  handleImageTransform: (action: ImageTransformAction) => void;
  /** Apply image position changes */
  handleApplyImagePosition: (data: ImagePositionData) => void;
  /** Open image properties dialog */
  handleOpenImageProperties: () => void;
  /** Apply image properties (alt text + border) */
  handleApplyImageProperties: (data: ImagePropertiesData) => void;
};

// ============================================================================
// HOOK
// ============================================================================

export const useImageHandlers = ({
  getActiveEditorView,
  focusActiveEditor,
  pmImageContext,
}: UseImageHandlersDeps): UseImageHandlersReturn => {
  // Image position dialog state
  const [imagePositionOpen, setImagePositionOpen] = useState(false);
  // Image properties dialog state
  const [imagePropsOpen, setImagePropsOpen] = useState(false);

  // Handle file selection for image insert
  const handleImageFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) {
        return;
      }

      const view = getActiveEditorView();
      if (!view) {
        return;
      }

      void insertImageFromFile(view, file, focusActiveEditor).catch(() => {
        // Image decode/read failures should not escape as unhandled rejections.
      });

      // Reset the input so the same file can be selected again
      e.target.value = "";
    },
    [getActiveEditorView, focusActiveEditor],
  );

  const handleImageWrapType = useCallback(
    (wrapType: string) => {
      const view = getActiveEditorView();
      if (!view || !pmImageContext) {
        return;
      }
      if (applyImageWrapType(view, pmImageContext.pos, wrapType)) {
        focusActiveEditor();
      }
    },
    [getActiveEditorView, focusActiveEditor, pmImageContext],
  );

  const handleImageTransform = useCallback(
    (action: ImageTransformAction) => {
      const view = getActiveEditorView();
      if (!view || !pmImageContext) {
        return;
      }
      if (applyImageTransform(view, pmImageContext.pos, action)) {
        focusActiveEditor();
      }
    },
    [getActiveEditorView, focusActiveEditor, pmImageContext],
  );

  const handleApplyImagePosition = useCallback(
    (data: ImagePositionData) => {
      const view = getActiveEditorView();
      if (!view || !pmImageContext) {
        return;
      }
      if (applyImagePosition(view, pmImageContext.pos, data)) {
        focusActiveEditor();
      }
    },
    [getActiveEditorView, focusActiveEditor, pmImageContext],
  );

  // Open image properties dialog
  const handleOpenImageProperties = useCallback(() => {
    setImagePropsOpen(true);
  }, []);

  const handleApplyImageProperties = useCallback(
    (data: ImagePropertiesData) => {
      const view = getActiveEditorView();
      if (!view || !pmImageContext) {
        return;
      }
      if (applyImageProperties(view, pmImageContext.pos, data)) {
        focusActiveEditor();
      }
    },
    [getActiveEditorView, focusActiveEditor, pmImageContext],
  );

  return {
    imagePositionOpen,
    setImagePositionOpen,
    imagePropsOpen,
    setImagePropsOpen,
    handleImageFileChange,
    handleImageWrapType,
    handleImageTransform,
    handleApplyImagePosition,
    handleOpenImageProperties,
    handleApplyImageProperties,
  };
};
