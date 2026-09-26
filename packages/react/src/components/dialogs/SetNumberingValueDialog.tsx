import { useEffect, useId, useState } from "react";
import { useTranslations } from "use-intl";

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

export type SetNumberingValueDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (value: number) => void;
};

/** `w:startOverride` is a non-negative decimal; the input stops far below its limit. */
const MIN_VALUE = 0;
const MAX_VALUE = 32767;
const DEFAULT_VALUE = 1;

/** *Set Numbering Value*: start the selected item's list over at a value. */
export function SetNumberingValueDialog({
  isOpen,
  onClose,
  onApply,
}: SetNumberingValueDialogProps) {
  const {
    Root: Dialog,
    Portal: DialogPortal,
    Backdrop: DialogBackdrop,
    Popup: DialogPopup,
    Title: DialogTitle,
    Close: DialogClose,
  } = useFolioUI().Dialog;
  const handleOpenChange = useCloseOnDialogOpenChange(onClose);
  const t = useTranslations("folio");
  const id = useId();
  // The raw field text, so the field can be cleared while a value is typed.
  const [text, setText] = useState(String(DEFAULT_VALUE));

  useEffect(() => {
    if (isOpen) {
      setText(String(DEFAULT_VALUE));
    }
  }, [isOpen]);

  const inputId = `${id}-numbering-value`;
  const parsed = text.trim() === "" ? Number.NaN : Number(text);
  const valid = Number.isFinite(parsed);
  const apply = () => {
    if (!valid) {
      return;
    }
    onApply(Math.min(MAX_VALUE, Math.max(MIN_VALUE, Math.trunc(parsed))));
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>
            {t("dialogs.setNumberingValue.title")}
          </DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <label className="flex flex-col gap-1" htmlFor={inputId}>
              <span className={DIALOG_LABEL_CLASS}>
                {t("dialogs.setNumberingValue.valueLabel")}
              </span>
              <input
                className={DIALOG_INPUT_CLASS}
                id={inputId}
                max={MAX_VALUE}
                min={MIN_VALUE}
                onChange={(event) => setText(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    event.preventDefault();
                    apply();
                  }
                }}
                type="number"
                value={text}
              />
            </label>
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>
              {t("common.cancel")}
            </DialogClose>
            <button
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!valid}
              onClick={apply}
              type="button"
            >
              {t("common.apply")}
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}
