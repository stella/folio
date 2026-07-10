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

export type HyperlinkBookmarkOption = {
  name: string;
  label?: string;
};

export type HyperlinkDialogData = {
  href: string;
  displayText: string;
  tooltip?: string;
};

export type HyperlinkDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onSubmit: (data: HyperlinkDialogData) => void;
  onRemove?: () => void;
  currentData?: Partial<HyperlinkDialogData>;
  selectedText?: string;
  bookmarks?: readonly HyperlinkBookmarkOption[];
};

type HyperlinkTargetType = "url" | "bookmark";

export function HyperlinkDialog({
  isOpen,
  onClose,
  onSubmit,
  onRemove,
  currentData,
  selectedText,
  bookmarks = [],
}: HyperlinkDialogProps) {
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
  const [targetType, setTargetType] = useState<HyperlinkTargetType>("url");
  const [url, setUrl] = useState("");
  const [bookmark, setBookmark] = useState("");
  const [displayText, setDisplayText] = useState("");
  const [tooltip, setTooltip] = useState("");
  const [touched, setTouched] = useState(false);

  useEffect(() => {
    if (!isOpen) {
      setTouched(false);
      return;
    }
    const href = currentData?.href ?? "";
    if (href.startsWith("#")) {
      setTargetType("bookmark");
      setBookmark(href.slice(1));
      setUrl("");
    } else {
      setTargetType("url");
      setUrl(href);
      setBookmark("");
    }
    setDisplayText(currentData?.displayText ?? selectedText ?? "");
    setTooltip(currentData?.tooltip ?? "");
  }, [isOpen, currentData, selectedText]);

  const urlError = targetType === "url" ? getUrlError(url) : "";
  const canSubmit =
    targetType === "bookmark"
      ? bookmark.trim().length > 0
      : url.trim().length > 0 && urlError.length === 0;
  const isEditing = Boolean(currentData?.href);
  const fieldIds = {
    targetType: `${id}-hyperlink-target-type`,
    url: `${id}-hyperlink-url`,
    bookmark: `${id}-hyperlink-bookmark`,
    displayText: `${id}-hyperlink-display-text`,
    tooltip: `${id}-hyperlink-tooltip`,
  };

  const handleSubmit = () => {
    setTouched(true);
    if (!canSubmit) {
      return;
    }
    const href = targetType === "bookmark" ? `#${bookmark.trim()}` : normalizeUrl(url);
    onSubmit({
      href,
      displayText: displayText.trim() || href,
      ...(tooltip.trim() ? { tooltip: tooltip.trim() } : {}),
    });
    onClose();
  };

  const handleRemove = () => {
    onRemove?.();
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>
            {isEditing ? "Edit Hyperlink" : "Insert Hyperlink"}
          </DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            {(bookmarks.length > 0 || targetType === "bookmark") && (
              <label className="flex flex-col gap-1" htmlFor={fieldIds.targetType}>
                <span className={DIALOG_LABEL_CLASS}>Link to</span>
                <select
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.targetType}
                  onChange={(e) => setTargetType(toTargetType(e.target.value))}
                  value={targetType}
                >
                  <option value="url">Web address</option>
                  <option value="bookmark">Bookmark</option>
                </select>
              </label>
            )}

            {targetType === "url" && (
              <label className="flex flex-col gap-1" htmlFor={fieldIds.url}>
                <span className={DIALOG_LABEL_CLASS}>Address</span>
                <input
                  aria-invalid={touched && urlError.length > 0}
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.url}
                  onBlur={() => setTouched(true)}
                  onChange={(e) => setUrl(e.target.value)}
                  placeholder="https://example.com"
                  value={url}
                />
                {touched && urlError.length > 0 && (
                  <span className="text-destructive text-xs">{urlError}</span>
                )}
              </label>
            )}

            {targetType === "bookmark" && (
              <label className="flex flex-col gap-1" htmlFor={fieldIds.bookmark}>
                <span className={DIALOG_LABEL_CLASS}>Bookmark</span>
                <select
                  className={DIALOG_INPUT_CLASS}
                  id={fieldIds.bookmark}
                  onChange={(e) => setBookmark(e.target.value)}
                  value={bookmark}
                >
                  <option value="">Select a bookmark</option>
                  {bookmarks.map((option) => (
                    <option key={option.name} value={option.name}>
                      {option.label ?? option.name}
                    </option>
                  ))}
                </select>
              </label>
            )}

            <label className="flex flex-col gap-1" htmlFor={fieldIds.displayText}>
              <span className={DIALOG_LABEL_CLASS}>Text to display</span>
              <input
                className={DIALOG_INPUT_CLASS}
                id={fieldIds.displayText}
                onChange={(e) => setDisplayText(e.target.value)}
                value={displayText}
              />
            </label>

            <label className="flex flex-col gap-1" htmlFor={fieldIds.tooltip}>
              <span className={DIALOG_LABEL_CLASS}>Screen tip</span>
              <input
                className={DIALOG_INPUT_CLASS}
                id={fieldIds.tooltip}
                onChange={(e) => setTooltip(e.target.value)}
                value={tooltip}
              />
            </label>
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            {isEditing && onRemove && (
              <button
                className="text-destructive me-auto rounded px-4 py-1.5 text-[13px]"
                onClick={handleRemove}
                type="button"
              >
                Remove
              </button>
            )}
            <DialogClose className={DIALOG_SECONDARY_BUTTON_CLASS}>Cancel</DialogClose>
            <button
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!canSubmit}
              onClick={handleSubmit}
              type="button"
            >
              {isEditing ? "Update" : "Insert"}
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}

function toTargetType(value: string): HyperlinkTargetType {
  return value === "bookmark" ? "bookmark" : "url";
}

function getUrlError(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return "Enter an address.";
  }
  if (/^(mailto:|tel:)/iu.test(trimmed)) {
    return trimmed.replace(/^(mailto:|tel:)/iu, "").length > 0 ? "" : "Enter an address.";
  }
  if (/^ftp:\/\//iu.test(trimmed)) {
    return trimmed.length > "ftp://".length ? "" : "Enter an address.";
  }
  try {
    const parsed = new URL(/^https?:\/\//iu.test(trimmed) ? trimmed : `https://${trimmed}`);
    return parsed.protocol === "http:" || parsed.protocol === "https:" ? "" : "Use a web address.";
  } catch {
    return "Enter a valid address.";
  }
}

function normalizeUrl(value: string): string {
  const trimmed = value.trim();
  if (/^(https?:\/\/|mailto:|tel:|ftp:\/\/)/iu.test(trimmed)) {
    return trimmed;
  }
  return `https://${trimmed}`;
}
