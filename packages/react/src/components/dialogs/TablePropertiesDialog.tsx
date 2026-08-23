/**
 * Table Properties Dialog — width type, width value, alignment.
 */

import { useCallback, useEffect, useId, useState } from "react";
import { useTranslations } from "use-intl";

import type { TablePropertiesCommand } from "@stll/folio-core/utils/tableOperations";

import { useFolioUI } from "../../ui/folio-ui";
import { useCloseOnDialogOpenChange } from "./dialogChrome";

export type TableProperties = TablePropertiesCommand;

export type TablePropertiesDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (props: TableProperties) => void;
  currentProps?: TableProperties;
};

const editableWidthType = (value: string): NonNullable<TableProperties["widthType"]> => {
  switch (value) {
    case "dxa":
    case "pct":
      return value;
    default:
      return "auto";
  }
};

const editableJustification = (value: string): NonNullable<TableProperties["justification"]> => {
  switch (value) {
    case "center":
    case "right":
      return value;
    default:
      return "left";
  }
};

export function TablePropertiesDialog({
  isOpen,
  onClose,
  onApply,
  currentProps,
}: TablePropertiesDialogProps) {
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
  const [width, setWidth] = useState(currentProps?.width ?? 0);
  const [widthType, setWidthType] = useState(editableWidthType(currentProps?.widthType ?? "auto"));
  const [justification, setJustification] = useState(
    editableJustification(currentProps?.justification ?? "left"),
  );

  useEffect(() => {
    if (isOpen) {
      setWidth(currentProps?.width ?? 0);
      setWidthType(editableWidthType(currentProps?.widthType ?? "auto"));
      setJustification(editableJustification(currentProps?.justification ?? "left"));
    }
  }, [isOpen, currentProps]);

  const handleApply = useCallback(() => {
    const justifValue =
      justification === "left" || justification === "center" || justification === "right"
        ? justification
        : ("left" as const);
    const props: TableProperties = {
      justification: justifValue,
    };
    if (widthType === "auto") {
      props.width = null;
      props.widthType = "auto";
    } else {
      props.width = width;
      props.widthType = widthType;
    }
    onApply(props);
    onClose();
  }, [width, widthType, justification, onApply, onClose]);

  const labelCls = "w-20 text-muted-foreground text-[13px]";
  const inputCls =
    "border-input bg-background text-foreground flex-1 rounded border px-2 py-1.5 text-[13px] outline-none";
  const fieldIds = {
    widthType: `${id}-tp-width-type`,
    width: `${id}-tp-width`,
    align: `${id}-tp-align`,
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className="fixed inset-0 z-[10000] bg-black/50" />
        <DialogPopup className="bg-popover fixed start-1/2 top-1/2 z-[10001] w-full max-w-[440px] min-w-[360px] -translate-x-1/2 -translate-y-1/2 rounded-lg border shadow-xl">
          <DialogTitle className="border-b px-5 py-3 text-base font-semibold">
            {t("dialogs.tableProperties.title")}
          </DialogTitle>

          <div className="flex flex-col gap-3 px-5 py-4">
            <div className="flex items-center gap-3">
              <label className={labelCls} htmlFor={fieldIds.widthType}>
                {t("dialogs.tableProperties.widthType")}
              </label>
              <select
                className={inputCls}
                id={fieldIds.widthType}
                onChange={(e) => setWidthType(editableWidthType(e.target.value))}
                value={widthType}
              >
                <option value="auto">{t("dialogs.tableProperties.widthTypes.auto")}</option>
                <option value="dxa">{t("dialogs.tableProperties.widthTypes.fixed")}</option>
                <option value="pct">{t("dialogs.tableProperties.widthTypes.percentage")}</option>
              </select>
            </div>

            {widthType !== "auto" && (
              <div className="flex items-center gap-3">
                <label className={labelCls} htmlFor={fieldIds.width}>
                  {t("dialogs.tableProperties.widthLabel")}
                </label>
                <input
                  className={inputCls}
                  id={fieldIds.width}
                  min={0}
                  onChange={(e) => setWidth(Number(e.target.value) || 0)}
                  step={widthType === "pct" ? 5 : 100}
                  type="number"
                  value={width}
                />
                <span className="text-muted-foreground text-[11px]">
                  {widthType === "pct"
                    ? t("dialogs.tableProperties.units.fiftiethsPercent")
                    : t("dialogs.tableProperties.units.twips")}
                </span>
              </div>
            )}

            <div className="flex items-center gap-3">
              <label className={labelCls} htmlFor={fieldIds.align}>
                {t("dialogs.tableProperties.alignmentLabel")}
              </label>
              <select
                className={inputCls}
                id={fieldIds.align}
                onChange={(e) => setJustification(editableJustification(e.target.value))}
                value={justification}
              >
                <option value="left">{t("dialogs.tableProperties.alignOptions.left")}</option>
                <option value="center">{t("dialogs.tableProperties.alignOptions.center")}</option>
                <option value="right">{t("dialogs.tableProperties.alignOptions.right")}</option>
              </select>
            </div>
          </div>

          <div className="flex justify-end gap-2 border-t px-5 py-3">
            <DialogClose className="border-input rounded border px-4 py-1.5 text-[13px]">
              {t("common.cancel")}
            </DialogClose>
            <button
              className="bg-primary text-primary-foreground rounded px-4 py-1.5 text-[13px] font-medium"
              onClick={handleApply}
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
