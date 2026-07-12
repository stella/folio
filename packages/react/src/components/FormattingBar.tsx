/**
 * FormattingBar — clean, minimal toolbar for legal document editing.
 *
 * Controls (left to right):
 * Undo Redo | Style ▾ | [host priorityExtra] | B I U [Format Painter] | Insert |
 * Font A▾ size Color | Alignment ¶ | Lists ◁ ▷ | ⋯ More (when collapsed) |
 * [host extras: ruler, zoom, ...]
 *
 * The font/color/alignment/list group collapses behind a "⋯ More" popover
 * whenever the bar is too narrow to show it inline (measured live, not a
 * fixed breakpoint); everything not shown here is still reachable via
 * keyboard shortcuts or host app chrome.
 */

import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, ReactNode } from "react";

import {
  BaselineIcon,
  BoldIcon,
  ChevronDownIcon,
  ImageIcon,
  ItalicIcon,
  MoreHorizontalIcon,
  PaintbrushIcon,
  PilcrowIcon,
  Redo2Icon,
  OmegaIcon,
  RulerIcon,
  SeparatorHorizontalIcon,
  TableIcon,
  TableOfContentsIcon,
  UnderlineIcon,
  Undo2Icon,
} from "lucide-react";
import { useTranslations } from "use-intl";

import type { ParagraphAlignment } from "@stll/folio-core/types/document";
import { cn } from "../lib/utils";
import { useFolioUI } from "../ui/folio-ui";
import type { ColorPreset } from "../ui/folio-ui";
import { containedHandler } from "../utils/contained-handler";
import { ToolbarButton, ToolbarGroup, ToolbarSeparator } from "./toolbarPrimitives";
import type { ToolbarProps, FormattingAction } from "./toolbarPrimitives";
import { AlignmentButtons } from "./ui/AlignmentButtons";
import { FontPicker } from "./ui/FontPicker";
import { normalizeFontFamilies } from "./ui/normalizeFontFamilies";
import { FontSizePicker } from "./ui/FontSizePicker";
import { ListButtons, createDefaultListState } from "./ui/ListButtons";
import { StylePicker } from "./ui/StylePicker";
import { ZoomControl } from "./ui/ZoomControl";

const ICON_SIZE = 16;
/** Buffer added to the measured content width before deciding it overflows,
 * absorbing sub-pixel rounding so the collapse decision does not flap. */
const OVERFLOW_EPSILON_PX = 1;

/** Default grid the Insert Table button requests (the toolbar has no size picker). */
const DEFAULT_TABLE_ROWS = 2;
const DEFAULT_TABLE_COLUMNS = 2;

/** Document color presets — hex values for OOXML compatibility. */
const DOCUMENT_COLOR_PRESETS: ColorPreset[] = [
  { label: "Black", value: "000000" },
  { label: "White", value: "FFFFFF" },
  { label: "Dark Red", value: "C00000" },
  { label: "Red", value: "FF0000" },
  { label: "Orange", value: "FFC000" },
  { label: "Yellow", value: "FFFF00" },
  { label: "Light Green", value: "92D050" },
  { label: "Green", value: "00B050" },
  { label: "Teal", value: "008080" },
  { label: "Light Blue", value: "00B0F0" },
  { label: "Blue", value: "0070C0" },
  { label: "Dark Blue", value: "002060" },
  { label: "Purple", value: "7030A0" },
  { label: "Dark Gray", value: "404040" },
  { label: "Medium Gray", value: "808080" },
  { label: "Light Gray", value: "C0C0C0" },
];

export type FormattingBarProps = {
  children?: ReactNode;
  /** Host controls that should stay in the primary row before text formatting */
  priorityExtra?: ReactNode;
  /** Extra controls rendered inline in the center column (after formatting buttons) */
  inlineExtra?: ReactNode;
  /** Display label for the style picker when the backing document is hydrating. */
  stylePickerLabel?: string | undefined;
  /** Computed preview style for the hydrating style picker label. */
  stylePickerLabelStyle?: CSSProperties | undefined;
  inline?: boolean;
} & ToolbarProps;

export function FormattingBar(props: FormattingBarProps) {
  const {
    currentFormatting = {},
    onFormat,
    onUndo,
    onRedo,
    canUndo = false,
    canRedo = false,
    disabled = false,
    className,
    style,
    enableShortcuts = true,
    editorRef,
    children,
    showStylePicker = true,
    showFormatPainter = true,
    formatPainterActive = false,
    onFormatPainter,
    showFontPicker = true,
    fontFamilies,
    showFontSizePicker = true,
    showTextColorPicker = true,
    showAlignmentButtons = true,
    showListButtons = true,
    showZoomControl,
    zoom,
    onZoomChange,
    rulerVisible = false,
    onToggleRuler,
    documentStyles,
    theme,
    onRefocusEditor,
    onInsertImage,
    onInsertTable,
    showTableInsert = true,
    onInsertPageBreak,
    onInsertTOC,
    onInsertSymbol,
    priorityExtra,
    inlineExtra,
    stylePickerLabel,
    stylePickerLabelStyle,
    inline = false,
  } = props;

  const folioUI = useFolioUI();
  const { Root: Menu, Trigger: MenuTrigger, Popup: MenuPopup } = folioUI.Menu;
  const ColorPicker = folioUI.ColorPicker;
  const t = useTranslations("folio");
  const barRef = useRef<HTMLDivElement>(null);
  // `scrollRef` is the horizontally-scrollable region holding the primary
  // controls (and the secondary group when it fits inline); `primaryRef`
  // wraps only the always-visible primary controls so its natural width can
  // be measured independent of whether the secondary group is shown.
  // `secondaryRef` wraps the inline separator plus secondary group while it is
  // rendered inline. Its last-measured width is cached in `secondaryWidthRef`,
  // so the collapse decision still has a number to compare against once the
  // group moves into the overflow popover (its width is effectively constant:
  // the group is all fixed-size icon controls, no localized text driving it).
  const scrollRef = useRef<HTMLDivElement>(null);
  const primaryRef = useRef<HTMLDivElement>(null);
  const secondaryRef = useRef<HTMLDivElement>(null);
  const secondaryWidthRef = useRef(0);
  // Start optimistic (assume the secondary group fits): the layout effect
  // below measures and corrects this before the first paint, so a narrow
  // initial width never flashes overflowing content.
  const [showSecondaryInline, setShowSecondaryInline] = useState(true);

  const handleFormat = useCallback(
    (action: FormattingAction) => {
      if (!disabled && onFormat) {
        onFormat(action);
      }
    },
    [disabled, onFormat],
  );
  const handleBold = useCallback(() => handleFormat("bold"), [handleFormat]);
  const handleItalic = useCallback(() => handleFormat("italic"), [handleFormat]);
  const handleUnderline = useCallback(() => handleFormat("underline"), [handleFormat]);

  const handleUndo = useCallback(() => {
    if (!disabled && canUndo && onUndo) {
      onUndo();
    }
  }, [disabled, canUndo, onUndo]);

  const handleRedo = useCallback(() => {
    if (!disabled && canRedo && onRedo) {
      onRedo();
    }
  }, [disabled, canRedo, onRedo]);

  // Format painter: a single click arms a one-shot paint, a double-click arms
  // the sticky (keep-on) mode. A real double-click also fires `click` first, so
  // native `click`/`dblclick` race on one button (arm → toggle-off → sticky) and
  // settle wrong; the one-shot arm is deferred instead, letting a second click
  // within the window pre-empt it. 300ms matches the platform double-click feel.
  const painterClickTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (painterClickTimer.current) {
        clearTimeout(painterClickTimer.current);
      }
    },
    [],
  );

  const handleFormatPainterClick = useCallback(() => {
    if (disabled || !onFormatPainter) {
      return;
    }
    if (painterClickTimer.current) {
      clearTimeout(painterClickTimer.current);
    }
    painterClickTimer.current = setTimeout(() => {
      painterClickTimer.current = null;
      onFormatPainter(false);
    }, 300);
  }, [disabled, onFormatPainter]);

  const handleFormatPainterDoubleClick = useCallback(() => {
    if (disabled || !onFormatPainter) {
      return;
    }
    if (painterClickTimer.current) {
      clearTimeout(painterClickTimer.current);
      painterClickTimer.current = null;
    }
    onFormatPainter(true);
  }, [disabled, onFormatPainter]);

  // Normalize the host's `fontFamilies` prop into the picker's FontOption[].
  // folio runs no React Compiler, so this memo is load-bearing: it keeps a
  // stable list reference across renders (undefined falls back to defaults).
  const fontPickerOptions = useMemo(() => normalizeFontFamilies(fontFamilies), [fontFamilies]);

  const handleFontFamilyChange = useCallback(
    (fontFamily: string) => {
      if (!disabled && onFormat) {
        onFormat({ type: "fontFamily", value: fontFamily });
        requestAnimationFrame(() => onRefocusEditor?.());
      }
    },
    [disabled, onFormat, onRefocusEditor],
  );

  const handleFontSizeChange = useCallback(
    (sizePt: number) => {
      if (!disabled && onFormat) {
        onFormat({ type: "fontSize", value: sizePt });
        requestAnimationFrame(() => onRefocusEditor?.());
      }
    },
    [disabled, onFormat, onRefocusEditor],
  );

  const handleTextColorSelect = useCallback(
    (hex: string) => {
      if (!disabled && onFormat) {
        onFormat({ type: "textColor", value: { rgb: hex } });
        requestAnimationFrame(() => onRefocusEditor?.());
      }
    },
    [disabled, onFormat, onRefocusEditor],
  );

  const handleTextColorClear = useCallback(() => {
    if (!disabled && onFormat) {
      onFormat({ type: "textColor", value: { auto: true } });
      requestAnimationFrame(() => onRefocusEditor?.());
    }
  }, [disabled, onFormat, onRefocusEditor]);

  const handleAlignmentChange = useCallback(
    (alignment: ParagraphAlignment) => {
      if (!disabled && onFormat) {
        onFormat({ type: "alignment", value: alignment });
      }
    },
    [disabled, onFormat],
  );

  const handleToggleDirection = useCallback(() => {
    if (!disabled && onFormat) {
      onFormat({ type: "toggleDirection" });
    }
  }, [disabled, onFormat]);

  const handleBulletList = useCallback(() => {
    if (!disabled && onFormat) {
      onFormat("bulletList");
    }
  }, [disabled, onFormat]);

  const handleNumberedList = useCallback(() => {
    if (!disabled && onFormat) {
      onFormat("numberedList");
    }
  }, [disabled, onFormat]);

  const handleIndent = useCallback(() => {
    if (!disabled && onFormat) {
      onFormat("indent");
    }
  }, [disabled, onFormat]);

  const handleOutdent = useCallback(() => {
    if (!disabled && onFormat) {
      onFormat("outdent");
    }
  }, [disabled, onFormat]);

  const handleStyleChange = useCallback(
    (styleId: string) => {
      if (!disabled && onFormat) {
        onFormat({ type: "applyStyle", value: styleId });
        requestAnimationFrame(() => onRefocusEditor?.());
      }
    },
    [disabled, onFormat, onRefocusEditor],
  );

  // After an insert, return focus to the editor so the caret lands in the new
  // content. Mouse clicks refocus via the bar's mouse-up handler, but keyboard
  // activation (Tab + Enter/Space) does not, so each insert refocuses itself.
  const handleInsertImage = useCallback(() => {
    if (!disabled && onInsertImage) {
      onInsertImage();
      requestAnimationFrame(() => onRefocusEditor?.());
    }
  }, [disabled, onInsertImage, onRefocusEditor]);

  const handleInsertTable = useCallback(() => {
    if (!disabled && onInsertTable) {
      onInsertTable(DEFAULT_TABLE_ROWS, DEFAULT_TABLE_COLUMNS);
      requestAnimationFrame(() => onRefocusEditor?.());
    }
  }, [disabled, onInsertTable, onRefocusEditor]);

  const handleInsertPageBreak = useCallback(() => {
    if (!disabled && onInsertPageBreak) {
      onInsertPageBreak();
      requestAnimationFrame(() => onRefocusEditor?.());
    }
  }, [disabled, onInsertPageBreak, onRefocusEditor]);

  const handleInsertTOC = useCallback(() => {
    if (!disabled && onInsertTOC) {
      onInsertTOC();
      requestAnimationFrame(() => onRefocusEditor?.());
    }
  }, [disabled, onInsertTOC, onRefocusEditor]);

  const handleInsertSymbol = useCallback(() => {
    if (!disabled && onInsertSymbol) {
      onInsertSymbol();
      requestAnimationFrame(() => onRefocusEditor?.());
    }
  }, [disabled, onInsertSymbol, onRefocusEditor]);

  const showTableButton = showTableInsert && Boolean(onInsertTable);
  const hasInsertControls =
    Boolean(onInsertImage) ||
    showTableButton ||
    Boolean(onInsertPageBreak) ||
    Boolean(onInsertTOC) ||
    Boolean(onInsertSymbol);

  // Keyboard shortcuts
  useEffect(() => {
    if (!enableShortcuts) {
      return;
    }

    const claimShortcut = (event: KeyboardEvent) => {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (!(event.target instanceof HTMLElement)) {
        return;
      }
      const target = event.target;
      const editorContainer = editorRef?.current;
      const barContainer = barRef.current;
      if (!editorContainer?.contains(target) && !barContainer?.contains(target)) {
        return;
      }

      const isCtrl = event.ctrlKey || event.metaKey;

      // Cmd+Enter — page break
      if (isCtrl && event.key === "Enter") {
        claimShortcut(event);
        handleFormat("insertPageBreak");
        return;
      }

      // Cmd+Shift shortcuts
      if (isCtrl && event.shiftKey) {
        switch (event.key) {
          case "=":
          case "+":
            claimShortcut(event);
            handleFormat("superscript");
            return;
          case "8":
            claimShortcut(event);
            handleBulletList();
            return;
          default:
            break;
        }
      }

      if (!isCtrl || event.altKey) {
        return;
      }

      switch (event.key.toLowerCase()) {
        case "b":
          claimShortcut(event);
          handleFormat("bold");
          break;
        case "i":
          claimShortcut(event);
          handleFormat("italic");
          break;
        case "u":
          claimShortcut(event);
          handleFormat("underline");
          break;
        case "d":
          claimShortcut(event);
          handleFormat("strikethrough");
          break;
        case "l":
          claimShortcut(event);
          handleAlignmentChange("left");
          break;
        case "e":
          claimShortcut(event);
          handleAlignmentChange("center");
          break;
        case "r":
          claimShortcut(event);
          handleAlignmentChange("right");
          break;
        case "j":
          claimShortcut(event);
          handleAlignmentChange("both");
          break;
        case "=":
          claimShortcut(event);
          handleFormat("subscript");
          break;
        case "]":
          claimShortcut(event);
          handleIndent();
          break;
        case "[":
          claimShortcut(event);
          handleOutdent();
          break;
        default:
          break;
      }
    };

    document.addEventListener("keydown", handleKeyDown, true);
    return () => document.removeEventListener("keydown", handleKeyDown, true);
  }, [
    enableShortcuts,
    handleFormat,
    handleAlignmentChange,
    handleBulletList,
    handleIndent,
    handleOutdent,
    editorRef,
  ]);

  // Decide whether the secondary group (font, color, alignment, lists) fits
  // inline by comparing the scroll container's real available width
  // (`clientWidth`, which already accounts for host `priorityExtra` and the
  // right-side host extras taking their share of the bar) against the
  // measured natural width of the primary controls plus the secondary
  // group. This replaces a fixed bar-width breakpoint, which ignored
  // however much room `priorityExtra` actually consumed and left a gap
  // between "too narrow to fit inline" and "narrow enough to show the ⋯
  // fallback" with no affordance in between.
  //
  // `useLayoutEffect` (not `useEffect`) so the very first measurement, and
  // any correction it makes to the optimistic initial state, happens
  // before the browser paints, avoiding a flash of overflowing content.
  useLayoutEffect(() => {
    if (inline) {
      setShowSecondaryInline(true);
      return undefined;
    }

    const scrollEl = scrollRef.current;
    const primaryEl = primaryRef.current;
    if (!scrollEl || !primaryEl) {
      return undefined;
    }

    const update = () => {
      const secondaryEl = secondaryRef.current;
      if (secondaryEl) {
        secondaryWidthRef.current = secondaryEl.offsetWidth;
      }
      const available = scrollEl.clientWidth;
      const scrollGap = Number.parseFloat(getComputedStyle(scrollEl).columnGap);
      const required =
        primaryEl.offsetWidth +
        secondaryWidthRef.current +
        (Number.isFinite(scrollGap) ? scrollGap : 0);
      const fits = required <= available + OVERFLOW_EPSILON_PX;
      setShowSecondaryInline((current) => (current === fits ? current : fits));
    };

    update();
    const observer = new ResizeObserver(update);
    observer.observe(scrollEl);
    observer.observe(primaryEl);
    if (secondaryRef.current) {
      observer.observe(secondaryRef.current);
    }

    return () => {
      observer.disconnect();
    };
  }, [inline, showSecondaryInline]);

  const handleBarMouseDown = useCallback((e: React.MouseEvent) => {
    if (!(e.target instanceof HTMLElement)) {
      e.preventDefault();
      return;
    }
    const target = e.target;
    if (target.tagName !== "INPUT" && target.tagName !== "SELECT") {
      e.preventDefault();
    }
  }, []);

  const handleBarMouseUp = useCallback(() => {
    requestAnimationFrame(() => onRefocusEditor?.());
  }, [onRefocusEditor]);

  const moreFormattingLabel = t("moreFormatting");
  const overflowTrigger = useMemo(
    () => (
      <button
        type="button"
        className="flex h-8 w-8 shrink-0 items-center justify-center rounded-md text-[var(--doc-text-muted)] transition-colors duration-100 hover:bg-[var(--doc-primary-light)] hover:text-[var(--doc-text)]"
        aria-label={moreFormattingLabel}
        title={moreFormattingLabel}
      />
    ),
    [moreFormattingLabel],
  );

  const secondaryControls = (
    <>
      <ToolbarGroup label={t("fontGroup")}>
        {showFontPicker && (
          <FontPicker
            value={currentFormatting.fontFamily || "Arial"}
            onChange={handleFontFamilyChange}
            fonts={fontPickerOptions}
            disabled={disabled}
            width={108}
            placeholder="Arial"
          />
        )}
        {showFontSizePicker && (
          <FontSizePicker
            value={
              currentFormatting.fontSize !== undefined ? currentFormatting.fontSize / 2 : undefined
            }
            onChange={handleFontSizeChange}
            disabled={disabled}
            ariaLabel={t("fontSize")}
          />
        )}
        {showTextColorPicker && (
          <ColorPicker
            presets={DOCUMENT_COLOR_PRESETS}
            columns={8}
            value={currentFormatting.color?.replace(/^#/u, "").toUpperCase()}
            onSelect={handleTextColorSelect}
            onClear={handleTextColorClear}
          >
            <ToolbarButton disabled={disabled} title={t("fontColor")} ariaLabel={t("fontColor")}>
              <div className="flex flex-col items-center gap-0">
                <BaselineIcon size={ICON_SIZE} />
                <span
                  className="mt-[-2px] h-1 w-4 rounded-sm"
                  style={{
                    backgroundColor: (() => {
                      if (currentFormatting.color) {
                        return (() => {
                          if (currentFormatting.color.startsWith("#")) {
                            return currentFormatting.color;
                          }
                          return `#${currentFormatting.color}`;
                        })();
                      }
                      return "#000000";
                    })(),
                  }}
                />
              </div>
              <ChevronDownIcon size={10} />
            </ToolbarButton>
          </ColorPicker>
        )}
      </ToolbarGroup>

      {showAlignmentButtons && (
        <ToolbarGroup label={t("alignmentGroup")}>
          <AlignmentButtons
            value={currentFormatting.alignment || "left"}
            onChange={handleAlignmentChange}
            disabled={disabled}
          />
          <ToolbarButton
            onClick={handleToggleDirection}
            active={currentFormatting.bidi}
            disabled={disabled}
            title={t("textDirection")}
            ariaLabel={t("textDirection")}
          >
            <PilcrowIcon size={ICON_SIZE} />
          </ToolbarButton>
        </ToolbarGroup>
      )}

      {showListButtons && (
        <ToolbarGroup label={t("listsGroup")}>
          <ListButtons
            listState={currentFormatting.listState || createDefaultListState()}
            onBulletList={handleBulletList}
            onNumberedList={handleNumberedList}
            onIndent={handleIndent}
            onOutdent={handleOutdent}
            disabled={disabled}
            showIndentButtons={true}
            compact
            hasIndent={(currentFormatting.indentLeft ?? 0) > 0}
          />
        </ToolbarGroup>
      )}

      {inlineExtra}
    </>
  );

  return (
    <div
      ref={barRef}
      className={cn(
        !inline &&
          "flex h-12 w-full items-center gap-0.5 overflow-hidden border-b border-[var(--doc-border)] bg-[var(--doc-page)] px-2 sm:px-4",
        className,
      )}
      style={inline ? { display: "contents", ...style } : style}
      role="toolbar"
      aria-label={t("formattingToolbar")}
      tabIndex={-1}
      data-folio-toolbar="true"
      onMouseDown={inline ? undefined : containedHandler(handleBarMouseDown)}
      onMouseUp={inline ? undefined : containedHandler(handleBarMouseUp)}
    >
      {/* Formatting controls: horizontally scrollable when the primary
          controls alone still overflow (the secondary group never scrolls —
          it either shows inline here or collapses into the ⋯ popover below). */}
      <div
        ref={scrollRef}
        className="flex min-w-0 flex-1 [scrollbar-width:none] items-center gap-0.5 overflow-x-auto overflow-y-hidden overscroll-x-contain [&::-webkit-scrollbar]:hidden"
      >
        {/* Primary controls: always visible; wrapped so their combined
            natural width can be measured independent of the secondary
            group (see the layout effect above). */}
        <div ref={primaryRef} className="flex shrink-0 items-center gap-0.5">
          {/* Undo / Redo */}
          <ToolbarGroup className="gap-0" label={t("historyGroup")}>
            <ToolbarButton
              className="h-8 w-7 rounded-e-none disabled:opacity-[0.35]"
              onClick={handleUndo}
              disabled={disabled || !canUndo}
              title={t("undoShortcut")}
              ariaLabel={t("undo")}
            >
              <Undo2Icon size={ICON_SIZE} />
            </ToolbarButton>
            <ToolbarButton
              className="h-8 w-7 rounded-s-none disabled:opacity-[0.35]"
              onClick={handleRedo}
              disabled={disabled || !canRedo}
              title={t("redoShortcut")}
              ariaLabel={t("redo")}
            >
              <Redo2Icon size={ICON_SIZE} />
            </ToolbarButton>
          </ToolbarGroup>

          <ToolbarSeparator />

          {/* Paragraph style gallery */}
          {showStylePicker && (
            <>
              <StylePicker
                value={currentFormatting.styleId || "Normal"}
                onChange={handleStyleChange}
                styles={documentStyles}
                theme={theme}
                disabled={disabled}
                displayLabel={stylePickerLabel}
                displayLabelStyle={stylePickerLabelStyle}
                className="shrink-0"
                width="clamp(112px, 15vw, 140px)"
              />
              <ToolbarSeparator />
            </>
          )}

          {priorityExtra && (
            <>
              {priorityExtra}
              <ToolbarSeparator />
            </>
          )}

          {/* Bold, Italic, Underline */}
          <ToolbarGroup label={t("textFormattingGroup")}>
            <ToolbarButton
              onClick={handleBold}
              active={currentFormatting.bold}
              disabled={disabled}
              title={t("boldShortcut")}
              ariaLabel={t("bold")}
            >
              <BoldIcon size={ICON_SIZE} />
            </ToolbarButton>
            <ToolbarButton
              onClick={handleItalic}
              active={currentFormatting.italic}
              disabled={disabled}
              title={t("italicShortcut")}
              ariaLabel={t("italic")}
            >
              <ItalicIcon size={ICON_SIZE} />
            </ToolbarButton>
            <ToolbarButton
              onClick={handleUnderline}
              active={currentFormatting.underline}
              disabled={disabled}
              title={t("underlineShortcut")}
              ariaLabel={t("underline")}
            >
              <UnderlineIcon size={ICON_SIZE} />
            </ToolbarButton>
            {showFormatPainter && onFormatPainter && (
              <ToolbarButton
                onClick={handleFormatPainterClick}
                onDoubleClick={handleFormatPainterDoubleClick}
                active={formatPainterActive}
                disabled={disabled}
                title={t("formatPainterShortcut")}
                ariaLabel={t("formatPainter")}
                testId="toolbar-format-painter"
              >
                <PaintbrushIcon size={ICON_SIZE} />
              </ToolbarButton>
            )}
          </ToolbarGroup>

          {/* Insert group — each control is opt-in: it renders only when its
            handler (and flag, for tables) is provided by the consumer. */}
          {hasInsertControls && (
            <>
              <ToolbarSeparator />
              <ToolbarGroup label={t("insertGroup")}>
                {onInsertImage && (
                  <ToolbarButton
                    onClick={handleInsertImage}
                    disabled={disabled}
                    title={t("insertImage")}
                    ariaLabel={t("insertImage")}
                  >
                    <ImageIcon size={ICON_SIZE} />
                  </ToolbarButton>
                )}
                {showTableButton && (
                  <ToolbarButton
                    onClick={handleInsertTable}
                    disabled={disabled}
                    title={t("insertTable")}
                    ariaLabel={t("insertTable")}
                    testId="toolbar-insert-table"
                  >
                    <TableIcon size={ICON_SIZE} />
                  </ToolbarButton>
                )}
                {onInsertPageBreak && (
                  <ToolbarButton
                    onClick={handleInsertPageBreak}
                    disabled={disabled}
                    title={t("insertPageBreak")}
                    ariaLabel={t("insertPageBreak")}
                  >
                    <SeparatorHorizontalIcon size={ICON_SIZE} />
                  </ToolbarButton>
                )}
                {onInsertTOC && (
                  <ToolbarButton
                    onClick={handleInsertTOC}
                    disabled={disabled}
                    title={t("insertTableOfContents")}
                    ariaLabel={t("insertTableOfContents")}
                  >
                    <TableOfContentsIcon size={ICON_SIZE} />
                  </ToolbarButton>
                )}
                {onInsertSymbol && (
                  <ToolbarButton
                    onClick={handleInsertSymbol}
                    disabled={disabled}
                    title={t("dialogs.insertSymbol.title")}
                    ariaLabel={t("dialogs.insertSymbol.title")}
                  >
                    <OmegaIcon size={ICON_SIZE} />
                  </ToolbarButton>
                )}
              </ToolbarGroup>
            </>
          )}
        </div>

        {showSecondaryInline && (
          <div ref={secondaryRef} className="flex shrink-0 items-center gap-0.5">
            <ToolbarSeparator />
            <div className="flex shrink-0 items-center gap-0.5">{secondaryControls}</div>
          </div>
        )}
      </div>

      {/* Overflow trigger for the secondary (font, color, alignment, lists)
          group: a sibling of the scrollable region above, not a child of it,
          so it is pinned in place and can never be scrolled out of view.
          Rendered only when that group has been measured to not fit inline. */}
      {!showSecondaryInline && (
        <>
          <ToolbarSeparator />
          <Menu>
            <MenuTrigger render={overflowTrigger}>
              <MoreHorizontalIcon size={ICON_SIZE} />
            </MenuTrigger>
            <MenuPopup align="end" className="max-w-[min(520px,calc(100vw-24px))]">
              <div className="flex w-[min(480px,calc(100vw-48px))] flex-wrap items-center gap-1 p-1">
                {secondaryControls}
              </div>
            </MenuPopup>
          </Menu>
        </>
      )}

      {/* Host extras (zoom, track changes, etc.) */}
      <div className="ms-auto flex shrink-0 items-center gap-1">
        {onToggleRuler && (
          <ToolbarGroup label={t("viewGroup")}>
            <ToolbarButton
              active={rulerVisible}
              onClick={onToggleRuler}
              title={t("ruler.toggle")}
              ariaLabel={t("ruler.toggle")}
              testId="toolbar-toggle-ruler"
            >
              <RulerIcon size={ICON_SIZE} />
            </ToolbarButton>
          </ToolbarGroup>
        )}
        {showZoomControl !== false && zoom !== undefined && onZoomChange !== undefined && (
          <ToolbarGroup label={t("zoomGroup")}>
            <ZoomControl value={zoom} onChange={onZoomChange} disabled={disabled} compact />
          </ToolbarGroup>
        )}
        {children}
      </div>
    </div>
  );
}
