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

export type SplitCellDialogData = {
  rows: number;
  columns: number;
  mergeBeforeSplit: boolean;
};

export type SplitCellDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onSplit: (data: SplitCellDialogData) => void;
  defaultRows?: number;
  defaultColumns?: number;
};

const MIN_PARTS = 1;
const MAX_PARTS = 63;

export function SplitCellDialog({
  isOpen,
  onClose,
  onSplit,
  defaultRows = 1,
  defaultColumns = 2,
}: SplitCellDialogProps) {
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
  const [rows, setRows] = useState(defaultRows);
  const [columns, setColumns] = useState(defaultColumns);
  const [mergeBeforeSplit, setMergeBeforeSplit] = useState(false);

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    setRows(defaultRows);
    setColumns(defaultColumns);
    setMergeBeforeSplit(false);
  }, [isOpen, defaultRows, defaultColumns]);

  const normalizedRows = clampInteger(rows);
  const normalizedColumns = clampInteger(columns);
  const canSplit = normalizedRows > 1 || normalizedColumns > 1;
  const fieldIds = {
    rows: `${id}-split-cell-rows`,
    columns: `${id}-split-cell-columns`,
    mergeBeforeSplit: `${id}-split-cell-merge-before-split`,
  };

  const handleSplit = () => {
    if (!canSplit) {
      return;
    }
    onSplit({ rows: normalizedRows, columns: normalizedColumns, mergeBeforeSplit });
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>Split Cell</DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <div className="grid grid-cols-2 gap-3">
              <label className="flex flex-col gap-1" htmlFor={fieldIds.columns}>
                <span className={DIALOG_LABEL_CLASS}>Columns</span>
                <input
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.columns}
                  max={MAX_PARTS}
                  min={MIN_PARTS}
                  onChange={(e) => setColumns(Number(e.target.value) || MIN_PARTS)}
                  type="number"
                  value={columns}
                />
              </label>
              <label className="flex flex-col gap-1" htmlFor={fieldIds.rows}>
                <span className={DIALOG_LABEL_CLASS}>Rows</span>
                <input
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.rows}
                  max={MAX_PARTS}
                  min={MIN_PARTS}
                  onChange={(e) => setRows(Number(e.target.value) || MIN_PARTS)}
                  type="number"
                  value={rows}
                />
              </label>
            </div>

            <label className="flex items-center gap-2" htmlFor={fieldIds.mergeBeforeSplit}>
              <input
                checked={mergeBeforeSplit}
                id={fieldIds.mergeBeforeSplit}
                onChange={(e) => setMergeBeforeSplit(e.target.checked)}
                type="checkbox"
              />
              <span className={DIALOG_LABEL_CLASS}>Merge selected cells before splitting</span>
            </label>
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>Cancel</DialogClose>
            <button
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!canSplit}
              onClick={handleSplit}
              type="button"
            >
              Split
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}

function clampInteger(value: number): number {
  return Math.min(MAX_PARTS, Math.max(MIN_PARTS, Math.trunc(value)));
}
