import { useEffect, useId, useState } from "react";

import { useFolioUI } from "../../ui/folio-ui";
import {
  DIALOG_BACKDROP_CLASS,
  DIALOG_BODY_CLASS,
  DIALOG_FOOTER_CLASS,
  DIALOG_LABEL_CLASS,
  DIALOG_POPUP_CLASS,
  DIALOG_PRIMARY_BUTTON_CLASS,
  DIALOG_SECONDARY_BUTTON_CLASS,
  DIALOG_TITLE_CLASS,
  useCloseOnDialogOpenChange,
} from "./dialogChrome";

export type PasteSpecialMode = "keepFormatting" | "mergeFormatting" | "plainText";

export type PasteSpecialDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onPaste: (mode: PasteSpecialMode) => void;
  defaultMode?: PasteSpecialMode;
};

const PASTE_MODES: Array<{ value: PasteSpecialMode; label: string; description: string }> = [
  {
    value: "keepFormatting",
    label: "Keep source formatting",
    description: "Preserve styles, links, and inline formatting from the clipboard.",
  },
  {
    value: "mergeFormatting",
    label: "Merge formatting",
    description: "Keep useful inline formatting while matching the destination paragraph.",
  },
  {
    value: "plainText",
    label: "Unformatted text",
    description: "Paste text only and use the destination formatting.",
  },
];

export function PasteSpecialDialog({
  isOpen,
  onClose,
  onPaste,
  defaultMode = "plainText",
}: PasteSpecialDialogProps) {
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
  const [mode, setMode] = useState<PasteSpecialMode>(defaultMode);

  useEffect(() => {
    if (isOpen) {
      setMode(defaultMode);
    }
  }, [isOpen, defaultMode]);

  const handlePaste = () => {
    onPaste(mode);
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>Paste Special</DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <fieldset className="flex flex-col gap-3">
              <legend className={DIALOG_LABEL_CLASS}>Paste mode</legend>
              {PASTE_MODES.map((option) => {
                const optionId = `${id}-paste-special-${option.value}`;
                return (
                  <label className="flex items-start gap-3" htmlFor={optionId} key={option.value}>
                    <input
                      checked={mode === option.value}
                      id={optionId}
                      onChange={() => setMode(option.value)}
                      type="radio"
                    />
                    <span className="flex flex-col gap-1">
                      <span className="text-foreground text-[13px] font-medium">
                        {option.label}
                      </span>
                      <span className="text-muted-foreground text-xs">{option.description}</span>
                    </span>
                  </label>
                );
              })}
            </fieldset>
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>Cancel</DialogClose>
            <button className={DIALOG_PRIMARY_BUTTON_CLASS} onClick={handlePaste} type="button">
              Paste
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}
