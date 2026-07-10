/**
 * HorizontalRuler Component — word-processor style
 *
 * 3 handles only:
 * - Left side: first-line indent (▼ down at top) + left indent (▲ up at bottom)
 * - Right side: right indent (▼ down at top)
 *
 * Margins shown as gray zones on the ruler edges.
 * Drag the boundary between gray/white to adjust page margins.
 * Drag tooltip shows value during any drag.
 */

import React, { useState, useRef, useCallback, useEffect, useMemo } from "react";
import type { CSSProperties } from "react";
import type { SectionProperties, TabStop } from "@stll/folio-core/types/document";
import { twipsToPixels, pixelsToTwips, formatPx } from "@stll/folio-core/utils/units";
import { useTranslations } from "use-intl";

// ============================================================================
// TYPES
// ============================================================================

export type HorizontalRulerProps = {
  sectionProps?: SectionProperties | null;
  zoom?: number;
  editable?: boolean;
  onLeftMarginChange?: (marginTwips: number) => void;
  onRightMarginChange?: (marginTwips: number) => void;
  onFirstLineIndentChange?: (indentTwips: number) => void;
  showFirstLineIndent?: boolean;
  firstLineIndent?: number;
  hangingIndent?: boolean;
  indentLeft?: number;
  indentRight?: number;
  onIndentLeftChange?: (indentTwips: number) => void;
  onIndentRightChange?: (indentTwips: number) => void;
  unit?: "inch" | "cm";
  className?: string;
  style?: CSSProperties;
  tabStops?: TabStop[] | null;
  onTabStopRemove?: (positionTwips: number) => void;
};

type MarkerType = "leftMargin" | "rightMargin" | "firstLineIndent" | "leftIndent" | "rightIndent";

// ============================================================================
// CONSTANTS
// ============================================================================

const DEFAULT_PAGE_WIDTH_TWIPS = 12240;
const DEFAULT_MARGIN_TWIPS = 1440;
const TWIPS_PER_INCH = 1440;
const TWIPS_PER_CM = 567;

const RULER_HEIGHT = 22;
const RULER_TEXT_COLOR = "var(--doc-text-muted)";
const RULER_TICK_COLOR = "var(--doc-text-subtle)";
const MARGIN_ZONE_COLOR = "var(--doc-shadow-subtle)";
const INDENT_COLOR = "var(--doc-primary)";
const INDENT_HOVER_COLOR = "var(--doc-primary-hover)";
const INDENT_ACTIVE_COLOR = "var(--doc-primary-hover)";

const TRI_SIZE = 5; // triangle half-width in px

function resolveIndentColor(isDragging: boolean, isHovered: boolean): string {
  if (isDragging) return INDENT_ACTIVE_COLOR;
  if (isHovered) return INDENT_HOVER_COLOR;
  return INDENT_COLOR;
}

// ============================================================================
// HELPERS
// ============================================================================

function formatValueForTooltip(twips: number, unit: "inch" | "cm"): string {
  if (unit === "inch") {
    return (twips / TWIPS_PER_INCH).toFixed(2) + '"';
  }
  return (twips / TWIPS_PER_CM).toFixed(1) + " cm";
}

// ============================================================================
// MAIN COMPONENT
// ============================================================================

export function HorizontalRuler({
  sectionProps,
  zoom = 1,
  editable = false,
  onLeftMarginChange,
  onRightMarginChange,
  onFirstLineIndentChange,
  showFirstLineIndent = false,
  firstLineIndent = 0,
  hangingIndent = false,
  indentLeft = 0,
  indentRight = 0,
  onIndentLeftChange,
  onIndentRightChange,
  unit = "inch",
  className = "",
  style,
  tabStops,
  onTabStopRemove,
}: HorizontalRulerProps): React.ReactElement {
  const t = useTranslations("folio");
  const [dragging, setDragging] = useState<MarkerType | null>(null);
  const [hoveredMarker, setHoveredMarker] = useState<MarkerType | null>(null);
  const [dragValue, setDragValue] = useState<number | null>(null);
  const [dragPositionPx, setDragPositionPx] = useState<number | null>(null);
  const rulerRef = useRef<HTMLDivElement>(null);

  // Page dimensions
  const pageWidthTwips = sectionProps?.pageWidth ?? DEFAULT_PAGE_WIDTH_TWIPS;
  const leftMarginTwips = sectionProps?.marginLeft ?? DEFAULT_MARGIN_TWIPS;
  const rightMarginTwips = sectionProps?.marginRight ?? DEFAULT_MARGIN_TWIPS;
  const contentTwips = pageWidthTwips - leftMarginTwips - rightMarginTwips;

  // Pixel conversions
  const pageWidthPx = twipsToPixels(pageWidthTwips) * zoom;
  const leftMarginPx = twipsToPixels(leftMarginTwips) * zoom;
  const rightMarginPx = twipsToPixels(rightMarginTwips) * zoom;
  const indentLeftPx = twipsToPixels(indentLeft) * zoom;
  const indentRightPx = twipsToPixels(indentRight) * zoom;

  // First line indent: hanging goes left, normal goes right
  const effectiveFirstLineIndent = hangingIndent ? -firstLineIndent : firstLineIndent;
  const firstLineIndentPx = twipsToPixels(effectiveFirstLineIndent) * zoom;

  // Handle positions (in px from ruler left edge)
  const leftIndentPosPx = leftMarginPx + indentLeftPx;
  const rightIndentPosPx = pageWidthPx - rightMarginPx - indentRightPx;
  const firstLinePosPx = leftMarginPx + indentLeftPx + firstLineIndentPx;

  // Values handleDrag reads change on every margin/indent update while a drag
  // is in flight. Stash them in a ref read inside the (stable) handler so the
  // mousemove/mouseup effect binds document listeners once per drag, not per frame.
  const dragParams = {
    dragging,
    zoom,
    pageWidthTwips,
    leftMarginTwips,
    rightMarginTwips,
    contentTwips,
    indentLeft,
    indentRight,
    onLeftMarginChange,
    onRightMarginChange,
    onFirstLineIndentChange,
    onIndentLeftChange,
    onIndentRightChange,
  };
  const dragParamsRef = useRef(dragParams);
  dragParamsRef.current = dragParams;

  const handleDragStart = useCallback(
    (e: React.MouseEvent, marker: MarkerType) => {
      if (!editable) return;
      e.preventDefault();
      e.stopPropagation();
      setDragging(marker);
    },
    [editable],
  );

  const handleDrag = useCallback((e: MouseEvent) => {
    const params = dragParamsRef.current;
    if (!params.dragging || !rulerRef.current) return;

    const rect = rulerRef.current.getBoundingClientRect();
    const x = e.clientX - rect.left;
    setDragPositionPx(x);
    const positionTwips = pixelsToTwips(x / params.zoom);

    if (params.dragging === "leftMargin") {
      const maxMargin = params.pageWidthTwips - params.rightMarginTwips - 720;
      const rounded = Math.round(Math.max(0, Math.min(positionTwips, maxMargin)));
      setDragValue(rounded);
      params.onLeftMarginChange?.(rounded);
    } else if (params.dragging === "rightMargin") {
      const fromRight = params.pageWidthTwips - positionTwips;
      const maxMargin = params.pageWidthTwips - params.leftMarginTwips - 720;
      const rounded = Math.round(Math.max(0, Math.min(fromRight, maxMargin)));
      setDragValue(rounded);
      params.onRightMarginChange?.(rounded);
    } else if (params.dragging === "firstLineIndent") {
      const base = params.leftMarginTwips + params.indentLeft;
      const indentFromBase = positionTwips - base;
      const maxIndent = params.contentTwips - params.indentLeft - params.indentRight - 720;
      const rounded = Math.round(Math.max(-params.indentLeft, Math.min(indentFromBase, maxIndent)));
      setDragValue(rounded);
      params.onFirstLineIndentChange?.(rounded);
    } else if (params.dragging === "leftIndent") {
      const indentFromMargin = positionTwips - params.leftMarginTwips;
      const maxIndent = params.contentTwips - params.indentRight - 720;
      const rounded = Math.round(Math.max(0, Math.min(indentFromMargin, maxIndent)));
      setDragValue(rounded);
      params.onIndentLeftChange?.(rounded);
    } else if (params.dragging === "rightIndent") {
      const rightEdge = params.pageWidthTwips - params.rightMarginTwips;
      const indentFromRight = rightEdge - positionTwips;
      const maxIndent = params.contentTwips - params.indentLeft - 720;
      const rounded = Math.round(Math.max(0, Math.min(indentFromRight, maxIndent)));
      setDragValue(rounded);
      params.onIndentRightChange?.(rounded);
    }
  }, []);

  const handleDragEnd = useCallback(() => {
    setDragging(null);
    setDragValue(null);
    setDragPositionPx(null);
  }, []);

  useEffect(() => {
    if (dragging) {
      document.addEventListener("mousemove", handleDrag);
      document.addEventListener("mouseup", handleDragEnd);
      return () => {
        document.removeEventListener("mousemove", handleDrag);
        document.removeEventListener("mouseup", handleDragEnd);
      };
    }
    return undefined;
  }, [dragging, handleDrag, handleDragEnd]);

  const ticks = useMemo(
    () => generateTicks(pageWidthTwips, zoom, unit),
    [pageWidthTwips, zoom, unit],
  );

  return (
    <div
      ref={rulerRef}
      className={`docx-horizontal-ruler ${className}`}
      style={{
        position: "relative",
        width: formatPx(pageWidthPx),
        height: RULER_HEIGHT,
        backgroundColor: "transparent",
        overflow: "visible",
        userSelect: "none",
        cursor: dragging ? "ew-resize" : "default",
        ...style,
      }}
      role="slider"
      aria-label={t("ruler.horizontal")}
      aria-valuemin={0}
      aria-valuemax={pageWidthTwips}
    >
      {/* Gray margin zones — click & drag anywhere in the gray area to adjust margin */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          width: formatPx(leftMarginPx),
          height: RULER_HEIGHT,
          backgroundColor: MARGIN_ZONE_COLOR,
          borderRight: "1px solid var(--doc-shadow-subtle)",
          cursor: editable ? "ew-resize" : "default",
          zIndex: 1,
        }}
        onMouseDown={
          editable && onLeftMarginChange ? (e) => handleDragStart(e, "leftMargin") : undefined
        }
      />
      <div
        style={{
          position: "absolute",
          top: 0,
          right: 0,
          width: formatPx(rightMarginPx),
          height: RULER_HEIGHT,
          backgroundColor: MARGIN_ZONE_COLOR,
          borderLeft: "1px solid var(--doc-shadow-subtle)",
          cursor: editable ? "ew-resize" : "default",
          zIndex: 1,
        }}
        onMouseDown={
          editable && onRightMarginChange ? (e) => handleDragStart(e, "rightMargin") : undefined
        }
      />

      {/* Tick marks */}
      <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>
        {ticks.map((tick, i) => (
          <RulerTick key={i} tick={tick} />
        ))}
      </div>

      {/* === 3 INDENT HANDLES (word-processor style) === */}

      {/* First-line indent — ▼ down triangle at top-left */}
      {showFirstLineIndent && (
        <IndentTriangle
          direction="down"
          positionPx={firstLinePosPx}
          editable={editable}
          isDragging={dragging === "firstLineIndent"}
          isHovered={hoveredMarker === "firstLineIndent"}
          marker="firstLineIndent"
          onDragStart={handleDragStart}
          onHoverChange={setHoveredMarker}
          label={t("ruler.firstLineIndent")}
        />
      )}

      {/* Left indent — ▲ up triangle at bottom-left */}
      {editable && onIndentLeftChange && (
        <IndentTriangle
          direction="up"
          positionPx={leftIndentPosPx}
          editable={editable}
          isDragging={dragging === "leftIndent"}
          isHovered={hoveredMarker === "leftIndent"}
          marker="leftIndent"
          onDragStart={handleDragStart}
          onHoverChange={setHoveredMarker}
          label={t("ruler.leftIndent")}
        />
      )}

      {/* Right indent — ▼ down triangle at top-right */}
      {editable && onIndentRightChange && (
        <IndentTriangle
          direction="down"
          positionPx={rightIndentPosPx}
          editable={editable}
          isDragging={dragging === "rightIndent"}
          isHovered={hoveredMarker === "rightIndent"}
          marker="rightIndent"
          onDragStart={handleDragStart}
          onHoverChange={setHoveredMarker}
          label={t("ruler.rightIndent")}
        />
      )}

      {/* Tab stop markers (display only). TabStop.position is measured from the
          text area's left edge, so offset by the left margin to place the marker
          on the physical page; removal still receives the raw text-area position. */}
      {tabStops?.map((tab) => (
        <TabStopMarker
          key={tab.position}
          tabStop={tab}
          positionPx={twipsToPixels(leftMarginTwips + tab.position) * zoom}
          onRemove={onTabStopRemove}
        />
      ))}

      {/* Drag tooltip */}
      {dragging && dragValue !== null && dragPositionPx !== null && (
        <DragTooltip value={formatValueForTooltip(dragValue, unit)} positionPx={dragPositionPx} />
      )}
    </div>
  );
}

// ============================================================================
// SUB-COMPONENTS
// ============================================================================

type TickData = {
  position: number;
  height: number;
  label?: string | undefined;
};

function RulerTick({ tick }: { tick: TickData }): React.ReactElement {
  return (
    <>
      <div
        style={{
          position: "absolute",
          left: formatPx(tick.position),
          bottom: 0,
          width: 1,
          height: tick.height,
          backgroundColor: RULER_TICK_COLOR,
        }}
      />
      {tick.label && (
        <div
          style={{
            position: "absolute",
            left: formatPx(tick.position),
            top: 3,
            transform: "translateX(-50%)",
            fontSize: "9px",
            color: RULER_TEXT_COLOR,
            fontFamily: "sans-serif",
            whiteSpace: "nowrap",
          }}
        >
          {tick.label}
        </div>
      )}
    </>
  );
}

/**
 * Indent triangle handle — word-processor style.
 * direction="down": ▼ anchored at top (first-line indent, right indent)
 * direction="up":   ▲ anchored at bottom (left indent)
 */
type IndentTriangleProps = {
  direction: "up" | "down";
  positionPx: number;
  editable: boolean;
  isDragging: boolean;
  isHovered: boolean;
  marker: MarkerType;
  onDragStart: (event: React.MouseEvent, marker: MarkerType) => void;
  onHoverChange: (marker: MarkerType | null) => void;
  label: string;
};

function IndentTriangle({
  direction,
  positionPx,
  editable,
  isDragging,
  isHovered,
  marker,
  onDragStart,
  onHoverChange,
  label,
}: IndentTriangleProps): React.ReactElement {
  const color = resolveIndentColor(isDragging, isHovered);
  const triHeight = Math.round(TRI_SIZE * 1.6);
  const handleMouseEnter = useCallback(() => onHoverChange(marker), [marker, onHoverChange]);
  const handleMouseLeave = useCallback(() => onHoverChange(null), [onHoverChange]);
  const handleMouseDown = useCallback(
    (event: React.MouseEvent) => onDragStart(event, marker),
    [marker, onDragStart],
  );

  const containerStyle: CSSProperties = {
    position: "absolute",
    left: formatPx(positionPx - TRI_SIZE),
    width: TRI_SIZE * 2,
    height: triHeight + 2,
    cursor: editable ? "ew-resize" : "default",
    zIndex: isDragging ? 10 : 4,
    ...(direction === "down" ? { top: 0 } : { bottom: 0 }),
  };

  const triangleStyle: CSSProperties =
    direction === "down"
      ? {
          position: "absolute",
          top: 1,
          left: 0,
          width: 0,
          height: 0,
          borderLeft: `${TRI_SIZE}px solid transparent`,
          borderRight: `${TRI_SIZE}px solid transparent`,
          borderTop: `${triHeight}px solid ${color}`,
          transition: "border-top-color 0.1s",
        }
      : {
          position: "absolute",
          bottom: 1,
          left: 0,
          width: 0,
          height: 0,
          borderLeft: `${TRI_SIZE}px solid transparent`,
          borderRight: `${TRI_SIZE}px solid transparent`,
          borderBottom: `${triHeight}px solid ${color}`,
          transition: "border-bottom-color 0.1s",
        };

  return (
    <div
      className="docx-ruler-indent"
      style={containerStyle}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      onMouseDown={handleMouseDown}
      role="slider"
      aria-label={label}
      aria-orientation="horizontal"
      tabIndex={editable ? 0 : -1}
    >
      <div style={triangleStyle} />
    </div>
  );
}

function DragTooltip({
  value,
  positionPx,
}: {
  value: string;
  positionPx: number;
}): React.ReactElement {
  return (
    <div
      style={{
        position: "absolute",
        left: formatPx(positionPx),
        top: -22,
        transform: "translateX(-50%)",
        backgroundColor: "var(--doc-text)",
        color: "var(--doc-on-primary)",
        fontSize: "10px",
        fontFamily: "sans-serif",
        padding: "2px 6px",
        borderRadius: 3,
        whiteSpace: "nowrap",
        pointerEvents: "none",
        zIndex: 20,
      }}
    >
      {value}
    </div>
  );
}

type TabStopMarkerProps = {
  tabStop: TabStop;
  positionPx: number;
  onRemove?: ((positionTwips: number) => void) | undefined;
};

const TAB_SYMBOLS: Record<string, string> = {
  left: "L",
  center: "C",
  right: "R",
  decimal: "D",
  bar: "|",
};

function TabStopMarker({ tabStop, positionPx, onRemove }: TabStopMarkerProps): React.ReactElement {
  const handleDoubleClick = useCallback(
    (event: React.MouseEvent) => {
      event.stopPropagation();
      onRemove?.(tabStop.position);
    },
    [onRemove, tabStop.position],
  );

  return (
    <div
      style={{
        position: "absolute",
        left: formatPx(positionPx - 5),
        bottom: 0,
        width: 10,
        height: 12,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: 8,
        fontWeight: 700,
        color: "var(--doc-text-muted)",
        cursor: "pointer",
        userSelect: "none",
      }}
      onDoubleClick={handleDoubleClick}
      title={`${tabStop.alignment} tab at ${(tabStop.position / 1440).toFixed(2)}"`}
    >
      {TAB_SYMBOLS[tabStop.alignment] || "L"}
    </div>
  );
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function generateTicks(pageWidthTwips: number, zoom: number, unit: "inch" | "cm"): TickData[] {
  const ticks: TickData[] = [];

  if (unit === "inch") {
    const eighthInchTwips = TWIPS_PER_INCH / 8;
    const totalEighths = Math.ceil(pageWidthTwips / eighthInchTwips);
    for (let i = 0; i <= totalEighths; i++) {
      const twipsPos = i * eighthInchTwips;
      if (twipsPos > pageWidthTwips) break;
      const pxPos = twipsToPixels(twipsPos) * zoom;
      if (i % 8 === 0) {
        ticks.push({ position: pxPos, height: 10, label: i / 8 > 0 ? String(i / 8) : undefined });
      } else if (i % 4 === 0) {
        ticks.push({ position: pxPos, height: 6 });
      } else if (i % 2 === 0) {
        ticks.push({ position: pxPos, height: 4 });
      } else {
        ticks.push({ position: pxPos, height: 2 });
      }
    }
  } else {
    const mmTwips = TWIPS_PER_CM / 10;
    const totalMm = Math.ceil(pageWidthTwips / mmTwips);
    for (let i = 0; i <= totalMm; i++) {
      const twipsPos = i * mmTwips;
      if (twipsPos > pageWidthTwips) break;
      const pxPos = twipsToPixels(twipsPos) * zoom;
      if (i % 10 === 0) {
        ticks.push({ position: pxPos, height: 10, label: i / 10 > 0 ? String(i / 10) : undefined });
      } else if (i % 5 === 0) {
        ticks.push({ position: pxPos, height: 6 });
      } else {
        ticks.push({ position: pxPos, height: 3 });
      }
    }
  }

  return ticks;
}

export function positionToMargin(
  positionPx: number,
  side: "left" | "right",
  pageWidthPx: number,
  zoom: number,
): number {
  const positionTwips = pixelsToTwips(positionPx / zoom);
  if (side === "left") return Math.max(0, positionTwips);
  return Math.max(0, pixelsToTwips(pageWidthPx / zoom) - positionTwips);
}

export function getRulerDimensions(
  sectionProps?: SectionProperties | null,
  zoom: number = 1,
): { width: number; leftMargin: number; rightMargin: number; contentWidth: number } {
  const pw = sectionProps?.pageWidth ?? DEFAULT_PAGE_WIDTH_TWIPS;
  const lm = sectionProps?.marginLeft ?? DEFAULT_MARGIN_TWIPS;
  const rm = sectionProps?.marginRight ?? DEFAULT_MARGIN_TWIPS;
  const width = twipsToPixels(pw) * zoom;
  const leftMargin = twipsToPixels(lm) * zoom;
  const rightMargin = twipsToPixels(rm) * zoom;
  return { width, leftMargin, rightMargin, contentWidth: width - leftMargin - rightMargin };
}

export function getMarginInUnits(marginTwips: number, unit: "inch" | "cm"): string {
  return unit === "inch"
    ? (marginTwips / TWIPS_PER_INCH).toFixed(2) + '"'
    : (marginTwips / TWIPS_PER_CM).toFixed(1) + " cm";
}

export function parseMarginFromUnits(value: string, unit: "inch" | "cm"): number | null {
  const num = parseFloat(value.replace(/[^\d.]/g, ""));
  if (isNaN(num)) return null;
  return Math.round(num * (unit === "inch" ? TWIPS_PER_INCH : TWIPS_PER_CM));
}

export default HorizontalRuler;
