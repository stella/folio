import { useCallback } from "react";

export function useCloseOnDialogOpenChange(onClose: () => void) {
  return useCallback(
    (open: boolean) => {
      if (!open) {
        onClose();
      }
    },
    [onClose],
  );
}

export const DIALOG_BACKDROP_CLASS = "fixed inset-0 z-[10000] bg-black/50";

export const DIALOG_POPUP_CLASS =
  "bg-popover fixed start-1/2 top-1/2 z-[10001] w-full max-w-[520px] min-w-[360px] -translate-x-1/2 -translate-y-1/2 rounded-lg border shadow-xl";

export const DIALOG_TITLE_CLASS = "border-b px-5 py-3 text-base font-semibold";

export const DIALOG_BODY_CLASS = "flex flex-col gap-4 px-5 py-4";

export const DIALOG_FOOTER_CLASS = "flex justify-end gap-2 border-t px-5 py-3";

export const DIALOG_LABEL_CLASS = "text-muted-foreground text-[13px]";

export const DIALOG_INPUT_CLASS =
  "border-input bg-background text-foreground rounded border px-2 py-1.5 text-[13px] outline-none";

export const DIALOG_SECONDARY_BUTTON_CLASS = "border-input rounded border px-4 py-1.5 text-[13px]";

export const DIALOG_PRIMARY_BUTTON_CLASS =
  "bg-primary text-primary-foreground rounded px-4 py-1.5 text-[13px] font-medium disabled:cursor-not-allowed disabled:opacity-50";
