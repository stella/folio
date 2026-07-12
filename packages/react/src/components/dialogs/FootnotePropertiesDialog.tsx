/**
 * Footnote & Endnote Properties Dialog
 *
 * Edits position, numbering format, start number, and restart rules.
 */

import { useId, useState } from "react";
import { useTranslations } from "use-intl";

import type {
  FootnoteProperties,
  EndnoteProperties,
  FootnotePosition,
  EndnotePosition,
  NoteNumberRestart,
  NumberFormat,
} from "@stll/folio-core/types/document";
import { useFolioUI } from "../../ui/folio-ui";
import { useCloseOnDialogOpenChange } from "./dialogChrome";

// ============================================================================
// TYPES
// ============================================================================

export type FootnotePropertiesDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (footnoteProps: FootnoteProperties, endnoteProps: EndnoteProperties) => void;
  footnotePr?: FootnoteProperties;
  endnotePr?: EndnoteProperties;
};

// ============================================================================
// NUMBER FORMAT OPTIONS
// ============================================================================

/** i18n message keys for the note number-format labels. */
type NumberFormatLabelKey =
  | "dialogs.footnoteProperties.formats.decimal"
  | "dialogs.footnoteProperties.formats.lowerRoman"
  | "dialogs.footnoteProperties.formats.upperRoman"
  | "dialogs.footnoteProperties.formats.lowerAlpha"
  | "dialogs.footnoteProperties.formats.upperAlpha"
  | "dialogs.footnoteProperties.formats.symbols";

const numberFormatOptions: { value: NumberFormat; labelKey: NumberFormatLabelKey }[] = [
  { value: "decimal", labelKey: "dialogs.footnoteProperties.formats.decimal" },
  { value: "lowerRoman", labelKey: "dialogs.footnoteProperties.formats.lowerRoman" },
  { value: "upperRoman", labelKey: "dialogs.footnoteProperties.formats.upperRoman" },
  { value: "lowerLetter", labelKey: "dialogs.footnoteProperties.formats.lowerAlpha" },
  { value: "upperLetter", labelKey: "dialogs.footnoteProperties.formats.upperAlpha" },
  { value: "chicago", labelKey: "dialogs.footnoteProperties.formats.symbols" },
];

// ============================================================================
// COMPONENT
// ============================================================================

export function FootnotePropertiesDialog({
  isOpen,
  onClose,
  onApply,
  footnotePr,
  endnotePr,
}: FootnotePropertiesDialogProps) {
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
  const [fnPosition, setFnPosition] = useState<FootnotePosition>(
    footnotePr?.position ?? "pageBottom",
  );
  const [fnNumFmt, setFnNumFmt] = useState<NumberFormat>(footnotePr?.numFmt ?? "decimal");
  const [fnNumStart, setFnNumStart] = useState<number>(footnotePr?.numStart ?? 1);
  const [fnRestart, setFnRestart] = useState<NoteNumberRestart>(
    footnotePr?.numRestart ?? "continuous",
  );

  const [enPosition, setEnPosition] = useState<EndnotePosition>(endnotePr?.position ?? "docEnd");
  const [enNumFmt, setEnNumFmt] = useState<NumberFormat>(endnotePr?.numFmt ?? "lowerRoman");
  const [enNumStart, setEnNumStart] = useState<number>(endnotePr?.numStart ?? 1);
  const [enRestart, setEnRestart] = useState<NoteNumberRestart>(
    endnotePr?.numRestart ?? "continuous",
  );

  const handleApply = () => {
    onApply(
      {
        position: fnPosition,
        numFmt: fnNumFmt,
        numStart: fnNumStart,
        numRestart: fnRestart,
      },
      {
        position: enPosition,
        numFmt: enNumFmt,
        numStart: enNumStart,
        numRestart: enRestart,
      },
    );
    onClose();
  };

  const labelCls = "block text-muted-foreground mb-1 text-xs";
  const selectCls =
    "border-input bg-background text-foreground mb-2 w-full rounded border px-2 py-1 text-[13px] outline-none";
  const inputCls =
    "border-input bg-background text-foreground w-[60px] rounded border px-2 py-1 text-[13px] outline-none";
  const sectionCls = "mb-4 rounded border p-3";
  const fieldIds = {
    fnPosition: `${id}-fn-position`,
    fnNumFmt: `${id}-fn-num-fmt`,
    fnStartAt: `${id}-fn-start-at`,
    fnNumbering: `${id}-fn-numbering`,
    enPosition: `${id}-en-position`,
    enNumFmt: `${id}-en-num-fmt`,
    enStartAt: `${id}-en-start-at`,
    enNumbering: `${id}-en-numbering`,
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className="fixed inset-0 z-[10000] bg-black/50" />
        <DialogPopup className="bg-popover fixed start-1/2 top-1/2 z-[10001] w-full max-w-[500px] min-w-[400px] -translate-x-1/2 -translate-y-1/2 rounded-lg border shadow-xl">
          <DialogTitle className="border-b px-5 py-3 text-base font-semibold">
            {t("dialogs.footnoteProperties.title")}
          </DialogTitle>

          <div className="flex flex-col gap-3 px-5 py-4">
            {/* Footnote section */}
            <div className={sectionCls}>
              <h4 className="mb-2 text-sm font-semibold">
                {t("dialogs.footnoteProperties.footnotes")}
              </h4>

              <label htmlFor={fieldIds.fnPosition} className={labelCls}>
                {t("dialogs.footnoteProperties.position")}
              </label>
              <select
                id={fieldIds.fnPosition}
                className={selectCls}
                value={fnPosition}
                onChange={(e) => setFnPosition(e.target.value as FootnotePosition)}
              >
                <option value="pageBottom">
                  {t("dialogs.footnoteProperties.footnotePositions.bottomOfPage")}
                </option>
                <option value="beneathText">
                  {t("dialogs.footnoteProperties.footnotePositions.belowText")}
                </option>
              </select>

              <label htmlFor={fieldIds.fnNumFmt} className={labelCls}>
                {t("dialogs.footnoteProperties.numberFormat")}
              </label>
              <select
                id={fieldIds.fnNumFmt}
                className={selectCls}
                value={fnNumFmt}
                onChange={(e) => setFnNumFmt(e.target.value as NumberFormat)}
              >
                {numberFormatOptions.map((o) => (
                  <option key={o.value} value={o.value}>
                    {t(o.labelKey)}
                  </option>
                ))}
              </select>

              <div className="flex items-center gap-3">
                <div>
                  <label htmlFor={fieldIds.fnStartAt} className={labelCls}>
                    {t("dialogs.footnoteProperties.startAt")}
                  </label>
                  <input
                    id={fieldIds.fnStartAt}
                    type="number"
                    min={1}
                    className={inputCls}
                    value={fnNumStart}
                    onChange={(e) => setFnNumStart(Number.parseInt(e.target.value, 10) || 1)}
                  />
                </div>
                <div className="flex-1">
                  <label htmlFor={fieldIds.fnNumbering} className={labelCls}>
                    {t("dialogs.footnoteProperties.numbering")}
                  </label>
                  <select
                    id={fieldIds.fnNumbering}
                    className={selectCls}
                    value={fnRestart}
                    onChange={(e) => setFnRestart(e.target.value as NoteNumberRestart)}
                  >
                    <option value="continuous">
                      {t("dialogs.footnoteProperties.numberingOptions.continuous")}
                    </option>
                    <option value="eachSect">
                      {t("dialogs.footnoteProperties.numberingOptions.restartSection")}
                    </option>
                    <option value="eachPage">
                      {t("dialogs.footnoteProperties.numberingOptions.restartPage")}
                    </option>
                  </select>
                </div>
              </div>
            </div>

            {/* Endnote section */}
            <div className={sectionCls}>
              <h4 className="mb-2 text-sm font-semibold">
                {t("dialogs.footnoteProperties.endnotes")}
              </h4>

              <label htmlFor={fieldIds.enPosition} className={labelCls}>
                {t("dialogs.footnoteProperties.position")}
              </label>
              <select
                id={fieldIds.enPosition}
                className={selectCls}
                value={enPosition}
                onChange={(e) => setEnPosition(e.target.value as EndnotePosition)}
              >
                <option value="docEnd">
                  {t("dialogs.footnoteProperties.endnotePositions.endOfDocument")}
                </option>
                <option value="sectEnd">
                  {t("dialogs.footnoteProperties.endnotePositions.endOfSection")}
                </option>
              </select>

              <label htmlFor={fieldIds.enNumFmt} className={labelCls}>
                {t("dialogs.footnoteProperties.numberFormat")}
              </label>
              <select
                id={fieldIds.enNumFmt}
                className={selectCls}
                value={enNumFmt}
                onChange={(e) => setEnNumFmt(e.target.value as NumberFormat)}
              >
                {numberFormatOptions.map((o) => (
                  <option key={o.value} value={o.value}>
                    {t(o.labelKey)}
                  </option>
                ))}
              </select>

              <div className="flex items-center gap-3">
                <div>
                  <label htmlFor={fieldIds.enStartAt} className={labelCls}>
                    {t("dialogs.footnoteProperties.startAt")}
                  </label>
                  <input
                    id={fieldIds.enStartAt}
                    type="number"
                    min={1}
                    className={inputCls}
                    value={enNumStart}
                    onChange={(e) => setEnNumStart(Number.parseInt(e.target.value, 10) || 1)}
                  />
                </div>
                <div className="flex-1">
                  <label htmlFor={fieldIds.enNumbering} className={labelCls}>
                    {t("dialogs.footnoteProperties.numbering")}
                  </label>
                  <select
                    id={fieldIds.enNumbering}
                    className={selectCls}
                    value={enRestart}
                    onChange={(e) => setEnRestart(e.target.value as NoteNumberRestart)}
                  >
                    <option value="continuous">
                      {t("dialogs.footnoteProperties.numberingOptions.continuous")}
                    </option>
                    <option value="eachSect">
                      {t("dialogs.footnoteProperties.numberingOptions.restartSection")}
                    </option>
                  </select>
                </div>
              </div>
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
