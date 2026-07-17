import { useEffect, useId, useState } from "react";
import { useTranslations } from "use-intl";

import { isAllowedExternalWatermarkImageUrl, type Watermark } from "@stll/folio-core/watermark";
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

export type WatermarkDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onApply: (watermark: Watermark | undefined) => void;
  currentWatermark?: Watermark;
};

type WatermarkMode = "none" | "text" | "picture";

const DEFAULT_TEXT_COLOR = "#C0C0C0";

export function WatermarkDialog({
  isOpen,
  onClose,
  onApply,
  currentWatermark,
}: WatermarkDialogProps) {
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
  const [mode, setMode] = useState<WatermarkMode>("text");
  const [text, setText] = useState("CONFIDENTIAL");
  const [font, setFont] = useState("Calibri");
  const [color, setColor] = useState(DEFAULT_TEXT_COLOR);
  const [diagonal, setDiagonal] = useState(true);
  const [opacityPercent, setOpacityPercent] = useState(50);
  const [imageRId, setImageRId] = useState("");
  const [imageTarget, setImageTarget] = useState("");
  const [imageTargetExternal, setImageTargetExternal] = useState(false);
  const [scalePercent, setScalePercent] = useState(100);
  const [washout, setWashout] = useState(true);

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    if (!currentWatermark) {
      setMode("text");
      setText("CONFIDENTIAL");
      setFont("Calibri");
      setColor(DEFAULT_TEXT_COLOR);
      setDiagonal(true);
      setOpacityPercent(50);
      setImageRId("");
      setImageTarget("");
      setImageTargetExternal(false);
      setScalePercent(100);
      setWashout(true);
      return;
    }
    setMode(currentWatermark.kind);
    if (currentWatermark.kind === "text") {
      setText(currentWatermark.text);
      setFont(currentWatermark.font ?? "Calibri");
      setColor(toColorInputValue(currentWatermark.color));
      setDiagonal(currentWatermark.diagonal ?? true);
      setOpacityPercent(Math.round((currentWatermark.opacity ?? 0.5) * 100));
      return;
    }
    setImageRId(currentWatermark.imageRId);
    setImageTarget(currentWatermark.imageTarget ?? "");
    setImageTargetExternal(currentWatermark.imageTargetExternal ?? false);
    setScalePercent(Math.round((currentWatermark.scale ?? 1) * 100));
    setWashout(currentWatermark.washout ?? true);
  }, [isOpen, currentWatermark]);

  const fieldIds = {
    mode: `${id}-watermark-mode`,
    text: `${id}-watermark-text`,
    font: `${id}-watermark-font`,
    color: `${id}-watermark-color`,
    diagonal: `${id}-watermark-diagonal`,
    opacity: `${id}-watermark-opacity`,
    imageRId: `${id}-watermark-image-rid`,
    imageTarget: `${id}-watermark-image-target`,
    imageTargetExternal: `${id}-watermark-image-external`,
    scale: `${id}-watermark-scale`,
    washout: `${id}-watermark-washout`,
  };

  // An external target becomes a `TargetMode="External"` relationship in the
  // saved package verbatim (see `docx/rezip.ts`); restrict it to http(s) so a
  // `file:` URL or UNC path can't ride along in the exported `.docx`.
  const imageTargetError =
    mode === "picture" && imageTargetExternal && imageTarget.trim().length > 0
      ? getExternalImageTargetError(imageTarget.trim())
      : "";

  const canApply =
    mode === "none" ||
    (mode === "text" && text.trim().length > 0) ||
    (mode === "picture" && imageRId.trim().length > 0 && imageTargetError.length === 0);

  const handleApply = () => {
    if (!canApply) {
      return;
    }
    if (mode === "none") {
      onApply(undefined);
      onClose();
      return;
    }
    if (mode === "text") {
      onApply({
        kind: "text",
        text: text.trim(),
        ...(font.trim() ? { font: font.trim() } : {}),
        color: stripHash(color),
        diagonal,
        opacity: clampPercent(opacityPercent) / 100,
      });
      onClose();
      return;
    }
    onApply({
      kind: "picture",
      imageRId: imageRId.trim(),
      ...(imageTarget.trim() ? { imageTarget: imageTarget.trim() } : {}),
      ...(imageTarget.trim() ? { imageTargetExternal } : {}),
      scale: clampScalePercent(scalePercent) / 100,
      washout,
    });
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>{t("dialogs.watermark.title")}</DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <label className="flex flex-col gap-1" htmlFor={fieldIds.mode}>
              <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.type")}</span>
              <select
                className={DIALOG_INPUT_CLASS}
                id={fieldIds.mode}
                onChange={(e) => setMode(toWatermarkMode(e.target.value))}
                value={mode}
              >
                <option value="text">{t("dialogs.watermark.textWatermark")}</option>
                <option value="picture">{t("dialogs.watermark.pictureWatermark")}</option>
                <option value="none">{t("dialogs.watermark.noWatermark")}</option>
              </select>
            </label>

            {mode === "text" && (
              <div className="grid grid-cols-2 gap-3">
                <label className="col-span-2 flex flex-col gap-1" htmlFor={fieldIds.text}>
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.text")}</span>
                  <input
                    className={DIALOG_INPUT_CLASS}
                    id={fieldIds.text}
                    onChange={(e) => setText(e.target.value)}
                    value={text}
                  />
                </label>
                <label className="flex flex-col gap-1" htmlFor={fieldIds.font}>
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.font")}</span>
                  <input
                    className={DIALOG_INPUT_CLASS}
                    id={fieldIds.font}
                    onChange={(e) => setFont(e.target.value)}
                    value={font}
                  />
                </label>
                <label className="flex flex-col gap-1" htmlFor={fieldIds.color}>
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.color")}</span>
                  <input
                    className={`${DIALOG_INPUT_CLASS} h-[34px]`}
                    id={fieldIds.color}
                    onChange={(e) => setColor(e.target.value)}
                    type="color"
                    value={color}
                  />
                </label>
                <label className="flex flex-col gap-1" htmlFor={fieldIds.opacity}>
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.opacity")}</span>
                  <input
                    className={DIALOG_INPUT_CLASS}
                    id={fieldIds.opacity}
                    max={100}
                    min={0}
                    onChange={(e) => setOpacityPercent(Number(e.target.value) || 0)}
                    type="number"
                    value={opacityPercent}
                  />
                </label>
                <label className="flex items-center gap-2 pt-6" htmlFor={fieldIds.diagonal}>
                  <input
                    checked={diagonal}
                    id={fieldIds.diagonal}
                    onChange={(e) => setDiagonal(e.target.checked)}
                    type="checkbox"
                  />
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.diagonal")}</span>
                </label>
              </div>
            )}

            {mode === "picture" && (
              <div className="grid grid-cols-2 gap-3">
                <label className="col-span-2 flex flex-col gap-1" htmlFor={fieldIds.imageRId}>
                  <span className={DIALOG_LABEL_CLASS}>
                    {t("dialogs.watermark.imageRelationshipId")}
                  </span>
                  <input
                    className={DIALOG_INPUT_CLASS}
                    id={fieldIds.imageRId}
                    onChange={(e) => setImageRId(e.target.value)}
                    placeholder="rId12"
                    value={imageRId}
                  />
                </label>
                <label className="col-span-2 flex flex-col gap-1" htmlFor={fieldIds.imageTarget}>
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.imageTarget")}</span>
                  <input
                    aria-invalid={imageTargetError.length > 0}
                    className={DIALOG_INPUT_CLASS}
                    id={fieldIds.imageTarget}
                    onChange={(e) => setImageTarget(e.target.value)}
                    placeholder="word/media/image1.png"
                    value={imageTarget}
                  />
                  {imageTargetError.length > 0 && (
                    <span className="text-destructive text-xs">{imageTargetError}</span>
                  )}
                </label>
                <label className="flex flex-col gap-1" htmlFor={fieldIds.scale}>
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.scale")}</span>
                  <input
                    className={DIALOG_INPUT_CLASS}
                    id={fieldIds.scale}
                    max={300}
                    min={1}
                    onChange={(e) => setScalePercent(Number(e.target.value) || 1)}
                    type="number"
                    value={scalePercent}
                  />
                </label>
                <label className="flex items-center gap-2 pt-6" htmlFor={fieldIds.washout}>
                  <input
                    checked={washout}
                    id={fieldIds.washout}
                    onChange={(e) => setWashout(e.target.checked)}
                    type="checkbox"
                  />
                  <span className={DIALOG_LABEL_CLASS}>{t("dialogs.watermark.washout")}</span>
                </label>
                <label
                  className="col-span-2 flex items-center gap-2"
                  htmlFor={fieldIds.imageTargetExternal}
                >
                  <input
                    checked={imageTargetExternal}
                    id={fieldIds.imageTargetExternal}
                    onChange={(e) => setImageTargetExternal(e.target.checked)}
                    type="checkbox"
                  />
                  <span className={DIALOG_LABEL_CLASS}>
                    {t("dialogs.watermark.targetIsExternalUrl")}
                  </span>
                </label>
              </div>
            )}
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>
              {t("common.cancel")}
            </DialogClose>
            <button
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!canApply}
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

function toWatermarkMode(value: string): WatermarkMode {
  if (value === "none" || value === "picture" || value === "text") {
    return value;
  }
  return "text";
}

function clampPercent(value: number): number {
  return Math.min(100, Math.max(0, value));
}

function clampScalePercent(value: number): number {
  return Math.min(300, Math.max(1, value));
}

function stripHash(value: string): string {
  return value.startsWith("#") ? value.slice(1) : value;
}

/** Mirrors `HyperlinkDialog.tsx`'s `getUrlError` — a plain-English, untranslated
 * field-level validation message (existing convention for this kind of inline
 * error in this dialog family). */
function getExternalImageTargetError(value: string): string {
  return isAllowedExternalWatermarkImageUrl(value) ? "" : "Use a web address (http/https).";
}

function toColorInputValue(value: string | undefined): string {
  if (!value || value === "auto") {
    return DEFAULT_TEXT_COLOR;
  }
  return value.startsWith("#") ? value : `#${value}`;
}
