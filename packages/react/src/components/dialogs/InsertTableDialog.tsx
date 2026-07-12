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

export type InsertTableDialogData = {
  rows: number;
  columns: number;
  autofit: boolean;
  styleId?: string;
};

export type InsertTableStyleOption = {
  id: string;
  name: string;
};

export type InsertTableDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onInsert: (data: InsertTableDialogData) => void;
  defaultRows?: number;
  defaultColumns?: number;
  styleOptions?: readonly InsertTableStyleOption[];
};

const GRID_ROWS = 8;
const GRID_COLUMNS = 10;
const MIN_ROWS = 1;
const MAX_ROWS = 100;
const MIN_COLUMNS = 1;
const MAX_COLUMNS = 20;
const DEFAULT_STYLE_OPTIONS: InsertTableStyleOption[] = [
  { id: "TableGrid", name: "Table Grid" },
  { id: "TableNormal", name: "Normal Table" },
  { id: "TableGridLight", name: "Grid Table Light" },
];

export function InsertTableDialog({
  isOpen,
  onClose,
  onInsert,
  defaultRows = 3,
  defaultColumns = 3,
  styleOptions = DEFAULT_STYLE_OPTIONS,
}: InsertTableDialogProps) {
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
  const [rows, setRows] = useState(defaultRows);
  const [columns, setColumns] = useState(defaultColumns);
  const [hoverRows, setHoverRows] = useState(0);
  const [hoverColumns, setHoverColumns] = useState(0);
  const [autofit, setAutofit] = useState(false);
  const [styleId, setStyleId] = useState("");

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    setRows(defaultRows);
    setColumns(defaultColumns);
    setHoverRows(0);
    setHoverColumns(0);
    setAutofit(false);
    setStyleId("");
  }, [isOpen, defaultRows, defaultColumns]);

  const normalizedRows = clampInteger(rows, MIN_ROWS, MAX_ROWS);
  const normalizedColumns = clampInteger(columns, MIN_COLUMNS, MAX_COLUMNS);
  const canInsert = Number.isFinite(rows) && Number.isFinite(columns);
  const fieldIds = {
    rows: `${id}-insert-table-rows`,
    columns: `${id}-insert-table-columns`,
    autofit: `${id}-insert-table-autofit`,
    style: `${id}-insert-table-style`,
  };

  const handleInsert = () => {
    if (!canInsert) {
      return;
    }
    onInsert({
      rows: normalizedRows,
      columns: normalizedColumns,
      autofit,
      ...(styleId ? { styleId } : {}),
    });
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>{t("dialogs.insertTable.title")}</DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <div
              className="grid w-max grid-cols-10 gap-1"
              onMouseLeave={() => {
                setHoverRows(0);
                setHoverColumns(0);
              }}
            >
              {Array.from({ length: GRID_ROWS }).map((_, rowIndex) =>
                Array.from({ length: GRID_COLUMNS }).map((__, columnIndex) => {
                  const row = rowIndex + 1;
                  const column = columnIndex + 1;
                  const active = row <= hoverRows && column <= hoverColumns;
                  return (
                    <button
                      aria-label={`${row} by ${column} table`}
                      className={`h-5 w-5 rounded-sm border ${
                        active ? "border-primary bg-primary/20" : "border-input bg-background"
                      }`}
                      key={`${row}-${column}`}
                      onClick={() => {
                        setRows(row);
                        setColumns(column);
                      }}
                      onMouseEnter={() => {
                        setHoverRows(row);
                        setHoverColumns(column);
                      }}
                      type="button"
                    />
                  );
                }),
              )}
            </div>
            <div className="text-muted-foreground text-xs">
              {hoverRows > 0
                ? `${hoverColumns} columns x ${hoverRows} rows`
                : `${normalizedColumns} columns x ${normalizedRows} rows`}
            </div>

            <div className="grid grid-cols-2 gap-3">
              <label className="flex flex-col gap-1" htmlFor={fieldIds.rows}>
                <span className={DIALOG_LABEL_CLASS}>{t("dialogs.insertTable.rowsLabel")}</span>
                <input
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.rows}
                  max={MAX_ROWS}
                  min={MIN_ROWS}
                  onChange={(e) => setRows(Number(e.target.value) || MIN_ROWS)}
                  type="number"
                  value={rows}
                />
              </label>
              <label className="flex flex-col gap-1" htmlFor={fieldIds.columns}>
                <span className={DIALOG_LABEL_CLASS}>{t("dialogs.insertTable.columnsLabel")}</span>
                <input
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.columns}
                  max={MAX_COLUMNS}
                  min={MIN_COLUMNS}
                  onChange={(e) => setColumns(Number(e.target.value) || MIN_COLUMNS)}
                  type="number"
                  value={columns}
                />
              </label>
            </div>

            <label className="flex items-center gap-2" htmlFor={fieldIds.autofit}>
              <input
                checked={autofit}
                id={fieldIds.autofit}
                onChange={(e) => setAutofit(e.target.checked)}
                type="checkbox"
              />
              <span className={DIALOG_LABEL_CLASS}>{t("dialogs.insertTable.autofit")}</span>
            </label>

            {styleOptions.length > 0 && (
              <label className="flex flex-col gap-1" htmlFor={fieldIds.style}>
                <span className={DIALOG_LABEL_CLASS}>{t("dialogs.insertTable.style")}</span>
                <select
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.style}
                  onChange={(e) => setStyleId(e.target.value)}
                  value={styleId}
                >
                  <option value="">{t("dialogs.insertTable.plainTable")}</option>
                  {styleOptions.map((style) => (
                    <option key={style.id} value={style.id}>
                      {style.name}
                    </option>
                  ))}
                </select>
              </label>
            )}
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>
              {t("common.cancel")}
            </DialogClose>
            <button
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!canInsert}
              onClick={handleInsert}
              type="button"
            >
              {t("common.insert")}
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}

function clampInteger(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, Math.trunc(value)));
}
