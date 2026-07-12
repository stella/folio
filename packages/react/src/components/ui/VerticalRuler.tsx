/**
 * VerticalRuler Component
 *
 * A vertical ruler that displays alongside the document with:
 * - Page height scale with tick marks
 * - Top and bottom margin indicators
 * - Optional dragging to adjust margins
 * - Support for zoom levels
 *
 * Similar to a word processor's vertical ruler.
 */

import React, { useState, useRef, useCallback, useEffect, useMemo } from "react";
import type { CSSProperties } from "react";
import type { SectionProperties } from "@stll/folio-core/types/document";
import { twipsToPixels, pixelsToTwips, formatPx } from "@stll/folio-core/utils/units";
import { useTranslations } from "use-intl";

// ============================================================================
// TYPES
// ============================================================================

export type VerticalRulerProps = {
  /** Section properties for page layout */
  sectionProps?: SectionProperties | null;
  /** Zoom level (1.0 = 100%) */
  zoom?: number;
  /** Whether margins can be dragged to adjust */
  editable?: boolean;
  /** Callback when top margin changes (in twips) */
  onTopMarginChange?: (marginTwips: number) => void;
  /** Callback when bottom margin changes (in twips) */
  onBottomMarginChange?: (marginTwips: number) => void;
  /** Unit to display (inches or cm) */
  unit?: "inch" | "cm";
  /** Additional CSS class name */
  className?: string;
  /** Additional inline styles */
  style?: CSSProperties;
};

type MarkerType = "topMargin" | "bottomMargin";

// ============================================================================
// CONSTANTS
// ============================================================================

const DEFAULT_PAGE_HEIGHT_TWIPS = 15840; // 11 inches
const DEFAULT_MARGIN_TWIPS = 1440; // 1 inch
const TWIPS_PER_INCH = 1440;
const TWIPS_PER_CM = 567;

// Ruler styling — word-processor style
// Exported so the outline toggle/panel can inset past the vertical ruler
// (it overlays the editor's left edge) instead of rendering on top of it.
export const RULER_WIDTH = 20;
const RULER_TEXT_COLOR = "var(--doc-text-muted)";
const RULER_TICK_COLOR = "var(--doc-text-subtle)";
const MARKER_COLOR = "var(--doc-primary)";
const MARKER_HOVER_COLOR = "var(--doc-primary)";
const MARKER_ACTIVE_COLOR = "var(--doc-primary-hover)";

function resolveMarkerColor(isDragging: boolean, isHovered: boolean): string {
  if (isDragging) return MARKER_ACTIVE_COLOR;
  if (isHovered) return MARKER_HOVER_COLOR;
  return MARKER_COLOR;
}

// ============================================================================
// MAIN COMPONENT
// ============================================================================

export function VerticalRuler({
  sectionProps,
  zoom = 1,
  editable = false,
  onTopMarginChange,
  onBottomMarginChange,
  unit = "inch",
  className = "",
  style,
}: VerticalRulerProps): React.ReactElement {
  const t = useTranslations("folio");
  const [dragging, setDragging] = useState<MarkerType | null>(null);
  const [hoveredMarker, setHoveredMarker] = useState<MarkerType | null>(null);
  const rulerRef = useRef<HTMLDivElement>(null);

  // Get page dimensions
  const pageHeightTwips = sectionProps?.pageHeight ?? DEFAULT_PAGE_HEIGHT_TWIPS;
  const topMarginTwips = sectionProps?.marginTop ?? DEFAULT_MARGIN_TWIPS;
  const bottomMarginTwips = sectionProps?.marginBottom ?? DEFAULT_MARGIN_TWIPS;

  // Convert to pixels with zoom
  const pageHeightPx = twipsToPixels(pageHeightTwips) * zoom;
  const topMarginPx = twipsToPixels(topMarginTwips) * zoom;
  const bottomMarginPx = twipsToPixels(bottomMarginTwips) * zoom;

  // Values handleDrag reads change on every margin update while a drag is in
  // flight. Stash them in a ref read inside the (stable) handler so the
  // mousemove/mouseup effect binds document listeners once per drag, not per frame.
  const dragParams = {
    dragging,
    zoom,
    pageHeightTwips,
    topMarginTwips,
    bottomMarginTwips,
    onTopMarginChange,
    onBottomMarginChange,
  };
  const dragParamsRef = useRef(dragParams);
  useEffect(() => {
    dragParamsRef.current = dragParams;
  }, [dragParams]);

  // Handle drag start
  const handleDragStart = useCallback(
    (e: React.MouseEvent, marker: MarkerType) => {
      if (!editable) return;
      e.preventDefault();
      setDragging(marker);
    },
    [editable],
  );

  // Handle drag
  const handleDrag = useCallback((e: MouseEvent) => {
    const params = dragParamsRef.current;
    if (!params.dragging || !rulerRef.current) return;

    const rect = rulerRef.current.getBoundingClientRect();
    const y = e.clientY - rect.top;

    const positionTwips = pixelsToTwips(y / params.zoom);

    if (params.dragging === "topMargin") {
      const maxMargin = params.pageHeightTwips - params.bottomMarginTwips - 720;
      const newMargin = Math.max(0, Math.min(positionTwips, maxMargin));
      params.onTopMarginChange?.(Math.round(newMargin));
    } else if (params.dragging === "bottomMargin") {
      const fromBottom = params.pageHeightTwips - positionTwips;
      const maxMargin = params.pageHeightTwips - params.topMarginTwips - 720;
      const newMargin = Math.max(0, Math.min(fromBottom, maxMargin));
      params.onBottomMarginChange?.(Math.round(newMargin));
    }
  }, []);

  // Handle drag end
  const handleDragEnd = useCallback(() => {
    setDragging(null);
  }, []);

  // Add/remove document event listeners
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

  // Generate tick marks
  const ticks = useMemo(
    () => generateVerticalTicks(pageHeightTwips, zoom, unit),
    [pageHeightTwips, zoom, unit],
  );

  const rulerStyle: CSSProperties = {
    position: "relative",
    width: RULER_WIDTH,
    height: formatPx(pageHeightPx),
    backgroundColor: "transparent",
    overflow: "visible",
    userSelect: "none",
    cursor: dragging ? "ns-resize" : "default",
    ...style,
  };

  return (
    <div
      ref={rulerRef}
      className={`docx-vertical-ruler ${className}`}
      style={rulerStyle}
      role="slider"
      aria-label={t("ruler.vertical")}
      aria-orientation="vertical"
    >
      {/* Tick marks */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          pointerEvents: "none",
        }}
      >
        {ticks.map((tick, index) => (
          <VerticalTick key={index} tick={tick} />
        ))}
      </div>

      {/* Top margin marker */}
      <VerticalMarginMarker
        type="topMargin"
        position={topMarginPx}
        editable={editable}
        isDragging={dragging === "topMargin"}
        isHovered={hoveredMarker === "topMargin"}
        onDragStart={handleDragStart}
        onHoverChange={setHoveredMarker}
      />

      {/* Bottom margin marker */}
      <VerticalMarginMarker
        type="bottomMargin"
        position={pageHeightPx - bottomMarginPx}
        editable={editable}
        isDragging={dragging === "bottomMargin"}
        isHovered={hoveredMarker === "bottomMargin"}
        onDragStart={handleDragStart}
        onHoverChange={setHoveredMarker}
      />
    </div>
  );
}

// ============================================================================
// SUB-COMPONENTS
// ============================================================================

type VerticalTickData = {
  position: number;
  width: number;
  label?: string | undefined;
};

function VerticalTick({ tick }: { tick: VerticalTickData }): React.ReactElement {
  const tickStyle: CSSProperties = {
    position: "absolute",
    top: formatPx(tick.position),
    right: 0,
    height: 1,
    width: tick.width,
    backgroundColor: RULER_TICK_COLOR,
  };

  const labelStyle: CSSProperties = {
    position: "absolute",
    top: formatPx(tick.position),
    left: 2,
    transform: "translateY(-50%)",
    fontSize: "9px",
    color: RULER_TEXT_COLOR,
    fontFamily: "sans-serif",
    whiteSpace: "nowrap",
  };

  return (
    <>
      <div style={tickStyle} />
      {tick.label && <div style={labelStyle}>{tick.label}</div>}
    </>
  );
}

type VerticalMarginMarkerProps = {
  type: "topMargin" | "bottomMargin";
  position: number;
  editable: boolean;
  isDragging: boolean;
  isHovered: boolean;
  onDragStart: (event: React.MouseEvent, marker: MarkerType) => void;
  onHoverChange: (marker: MarkerType | null) => void;
};

function VerticalMarginMarker({
  type,
  position,
  editable,
  isDragging,
  isHovered,
  onDragStart,
  onHoverChange,
}: VerticalMarginMarkerProps): React.ReactElement {
  const t = useTranslations("folio");
  const color = resolveMarkerColor(isDragging, isHovered);
  const handleMouseEnter = useCallback(() => onHoverChange(type), [onHoverChange, type]);
  const handleMouseLeave = useCallback(() => onHoverChange(null), [onHoverChange]);
  const handleMouseDown = useCallback(
    (event: React.MouseEvent) => onDragStart(event, type),
    [onDragStart, type],
  );

  const markerStyle: CSSProperties = {
    position: "absolute",
    top: formatPx(position - 5),
    right: 0,
    width: RULER_WIDTH,
    height: 10,
    cursor: editable ? "ns-resize" : "default",
    zIndex: isDragging ? 10 : 1,
  };

  // Triangle pointing left (for top) or right (for bottom)
  const triangleStyle: CSSProperties = {
    position: "absolute",
    top: 0,
    right: 2,
    width: 0,
    height: 0,
    borderTop: "5px solid transparent",
    borderBottom: "5px solid transparent",
    borderRight: `8px solid ${color}`,
    transition: "border-right-color 0.1s",
  };

  return (
    <div
      className={`docx-ruler-marker docx-ruler-marker-${type}`}
      style={markerStyle}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      onMouseDown={handleMouseDown}
      role="slider"
      aria-label={type === "topMargin" ? t("ruler.topMargin") : t("ruler.bottomMargin")}
      aria-orientation="vertical"
      tabIndex={editable ? 0 : -1}
    >
      <div style={triangleStyle} />
    </div>
  );
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function generateVerticalTicks(
  pageHeightTwips: number,
  zoom: number,
  unit: "inch" | "cm",
): VerticalTickData[] {
  const ticks: VerticalTickData[] = [];

  if (unit === "inch") {
    const eighthInchTwips = TWIPS_PER_INCH / 8;
    const totalEighths = Math.ceil(pageHeightTwips / eighthInchTwips);

    for (let i = 0; i <= totalEighths; i++) {
      const twipsPos = i * eighthInchTwips;
      if (twipsPos > pageHeightTwips) break;

      const pxPos = twipsToPixels(twipsPos) * zoom;

      if (i % 8 === 0) {
        const inches = i / 8;
        ticks.push({
          position: pxPos,
          width: 10,
          label: inches > 0 ? String(inches) : undefined,
        });
      } else if (i % 4 === 0) {
        ticks.push({ position: pxPos, width: 6 });
      } else if (i % 2 === 0) {
        ticks.push({ position: pxPos, width: 4 });
      } else {
        ticks.push({ position: pxPos, width: 2 });
      }
    }
  } else {
    const mmTwips = TWIPS_PER_CM / 10;
    const totalMm = Math.ceil(pageHeightTwips / mmTwips);

    for (let i = 0; i <= totalMm; i++) {
      const twipsPos = i * mmTwips;
      if (twipsPos > pageHeightTwips) break;

      const pxPos = twipsToPixels(twipsPos) * zoom;

      if (i % 10 === 0) {
        const cm = i / 10;
        ticks.push({
          position: pxPos,
          width: 10,
          label: cm > 0 ? String(cm) : undefined,
        });
      } else if (i % 5 === 0) {
        ticks.push({ position: pxPos, width: 6 });
      } else {
        ticks.push({ position: pxPos, width: 3 });
      }
    }
  }

  return ticks;
}

export default VerticalRuler;
