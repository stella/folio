import { useEffect, useId, useState } from "react";

import { useFolioUI } from "../../ui/folio-ui";
import {
  DIALOG_BACKDROP_CLASS,
  DIALOG_BODY_CLASS,
  DIALOG_FOOTER_CLASS,
  DIALOG_INPUT_CLASS,
  DIALOG_LABEL_CLASS,
  DIALOG_POPUP_CLASS,
  DIALOG_PRIMARY_BUTTON_CLASS,
  DIALOG_SECONDARY_BUTTON_CLASS,
  DIALOG_TITLE_CLASS,
  useCloseOnDialogOpenChange,
} from "./dialogChrome";

export type InsertImageDialogData = {
  file: File;
  alt?: string;
  width?: number;
  height?: number;
};

export type InsertImageDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onInsert: (data: InsertImageDialogData) => void;
  accept?: string;
};

export function InsertImageDialog({
  isOpen,
  onClose,
  onInsert,
  accept = "image/png,image/jpeg,image/gif,image/webp",
}: InsertImageDialogProps) {
  const {
    Root: Dialog,
    Portal: DialogPortal,
    Backdrop: DialogBackdrop,
    Popup: DialogPopup,
    Title: DialogTitle,
    Close: DialogClose,
  } = useFolioUI().Dialog;
  const handleOpenChange = useCloseOnDialogOpenChange(onClose);
  const id = useId();
  const [file, setFile] = useState<File | null>(null);
  const [alt, setAlt] = useState("");
  const [width, setWidth] = useState("");
  const [height, setHeight] = useState("");

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    setFile(null);
    setAlt("");
    setWidth("");
    setHeight("");
  }, [isOpen]);

  const fieldIds = {
    file: `${id}-insert-image-file`,
    alt: `${id}-insert-image-alt`,
    width: `${id}-insert-image-width`,
    height: `${id}-insert-image-height`,
  };

  const handleInsert = () => {
    if (!file) {
      return;
    }
    const parsedWidth = toPositiveNumber(width);
    const parsedHeight = toPositiveNumber(height);
    onInsert({
      file,
      ...(alt.trim() ? { alt: alt.trim() } : {}),
      ...(parsedWidth ? { width: parsedWidth } : {}),
      ...(parsedHeight ? { height: parsedHeight } : {}),
    });
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>Insert Image</DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <label className="flex flex-col gap-1" htmlFor={fieldIds.file}>
              <span className={DIALOG_LABEL_CLASS}>Image file</span>
              <input
                accept={accept}
                className={DIALOG_INPUT_CLASS}
                id={fieldIds.file}
                onChange={(e) => setFile(e.target.files?.[0] ?? null)}
                type="file"
              />
            </label>

            <label className="flex flex-col gap-1" htmlFor={fieldIds.alt}>
              <span className={DIALOG_LABEL_CLASS}>Alt text</span>
              <input
                className={DIALOG_INPUT_CLASS}
                id={fieldIds.alt}
                onChange={(e) => setAlt(e.target.value)}
                placeholder={file?.name ?? ""}
                value={alt}
              />
            </label>

            <div className="grid grid-cols-2 gap-3">
              <label className="flex flex-col gap-1" htmlFor={fieldIds.width}>
                <span className={DIALOG_LABEL_CLASS}>Width</span>
                <input
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.width}
                  min={1}
                  onChange={(e) => setWidth(e.target.value)}
                  placeholder="Auto"
                  type="number"
                  value={width}
                />
              </label>
              <label className="flex flex-col gap-1" htmlFor={fieldIds.height}>
                <span className={DIALOG_LABEL_CLASS}>Height</span>
                <input
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.height}
                  min={1}
                  onChange={(e) => setHeight(e.target.value)}
                  placeholder="Auto"
                  type="number"
                  value={height}
                />
              </label>
            </div>
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>Cancel</DialogClose>
            <button
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!file}
              onClick={handleInsert}
              type="button"
            >
              Insert
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}

function toPositiveNumber(value: string): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}
