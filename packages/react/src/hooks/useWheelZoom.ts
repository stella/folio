/**
 * useWheelZoom Hook
 *
 * Enables Ctrl+scroll (or Cmd+scroll on Mac) to zoom in/out.
 * Features:
 * - Configurable zoom range and step
 * - Smooth zoom transitions
 * - Pinch-to-zoom support on trackpads (browsers report pinch as a
 *   `ctrlKey` wheel event)
 * - Zoom reset (Ctrl+0)
 * - Zoom in/out shortcuts (Ctrl++, Ctrl+-)
 *
 * The hook is uncontrolled by default (owns its own zoom state). Pass
 * `getCurrentZoom` to run it controlled: deltas then compose against an external
 * zoom owner (e.g. DocxEditor's viewport-anchored zoom) and the hook only emits
 * through `onZoomChange`, so the two never diverge.
 */

import { useCallback, useEffect, useRef, useState } from "react";
import type { RefObject } from "react";
import { ZOOM_MAX, ZOOM_MIN, ZOOM_STEP } from "@stll/folio-core/utils/zoom";

// ============================================================================
// TYPES
// ============================================================================

/**
 * Options for useWheelZoom hook
 */
export type UseWheelZoomOptions = {
  /** Initial zoom level (default: 1.0) */
  initialZoom?: number;
  /** Minimum zoom level (default: 0.25) */
  minZoom?: number;
  /** Maximum zoom level (default: 4.0) */
  maxZoom?: number;
  /** Zoom step for each scroll event (default: 0.1) */
  zoomStep?: number;
  /** Whether zoom is enabled (default: true) */
  enabled?: boolean;
  /** Container element ref to attach wheel listener */
  containerRef?: RefObject<HTMLElement | null>;
  /**
   * Live current zoom getter. When provided the hook runs controlled: it reads
   * the current zoom from here (instead of its own state) when computing deltas
   * and only emits the next value through `onZoomChange`. Lets the hook compose
   * with an external zoom owner without a second, divergent source of truth.
   */
  getCurrentZoom?: () => number;
  /** Callback when zoom changes */
  onZoomChange?: (zoom: number) => void;
  /** Whether to enable keyboard shortcuts (Ctrl++, Ctrl+-, Ctrl+0) */
  enableKeyboardShortcuts?: boolean;
  /** Whether to prevent default browser zoom behavior */
  preventDefault?: boolean;
};

/**
 * Return value of useWheelZoom hook
 */
export type UseWheelZoomReturn = {
  /** Current zoom level */
  zoom: number;
  /** Set zoom level directly */
  setZoom: (zoom: number) => void;
  /** Zoom in by step */
  zoomIn: () => void;
  /** Zoom out by step */
  zoomOut: () => void;
  /** Reset zoom to initial level */
  resetZoom: () => void;
  /** Reset zoom to 100% */
  zoomTo100: () => void;
  /** Zoom to fit width */
  zoomToFit: (containerWidth: number, contentWidth: number) => void;
  /** Whether currently at minimum zoom */
  isMinZoom: boolean;
  /** Whether currently at maximum zoom */
  isMaxZoom: boolean;
  /** Zoom percentage (e.g., 100 for zoom level 1.0) */
  zoomPercent: number;
  /** Wheel event handler (for manual attachment) */
  handleWheel: (event: WheelEvent) => void;
  /** Keyboard event handler (for manual attachment) */
  handleKeyDown: (event: KeyboardEvent) => void;
};

// ============================================================================
// CONSTANTS
// ============================================================================

const DEFAULT_INITIAL_ZOOM = 1;
const DEFAULT_MIN_ZOOM = ZOOM_MIN;
const DEFAULT_MAX_ZOOM = ZOOM_MAX;
const DEFAULT_ZOOM_STEP = ZOOM_STEP;

/**
 * Per-event zoom sensitivity for Ctrl/Cmd+wheel and trackpad pinch. Multiplies
 * deltaY inside an exponential so zoom scales continuously: small pinch deltas
 * nudge gently while a coarse mouse-wheel notch still moves a perceptible amount.
 */
const WHEEL_ZOOM_SENSITIVITY = 0.003;

/**
 * Preset zoom levels for wheel/keyboard snapping within the shared
 * `[ZOOM_MIN, ZOOM_MAX]` range (see `@stll/folio-core/utils/zoom`).
 */
export const ZOOM_PRESETS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4];

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Clamp a value between min and max
 */
function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

/**
 * Round zoom to 2 decimal places
 */
function roundZoom(zoom: number): number {
  return Math.round(zoom * 100) / 100;
}

/**
 * Find nearest preset zoom level
 */
function nearestPreset(zoom: number): number {
  // SAFETY: ZOOM_PRESETS is a non-empty constant array
  let nearest = ZOOM_PRESETS[0]!;
  let minDiff = Math.abs(zoom - nearest);

  for (const preset of ZOOM_PRESETS) {
    const diff = Math.abs(zoom - preset);
    if (diff < minDiff) {
      minDiff = diff;
      nearest = preset;
    }
  }

  return nearest;
}

/**
 * Get next preset zoom level (for zoom in)
 */
function nextPreset(currentZoom: number): number {
  for (const preset of ZOOM_PRESETS) {
    if (preset > currentZoom + 0.01) {
      return preset;
    }
  }
  // oxlint-disable-next-line typescript/no-non-null-assertion
  return ZOOM_PRESETS.at(-1)!;
}

/**
 * Get previous preset zoom level (for zoom out)
 */
function prevPreset(currentZoom: number): number {
  for (let i = ZOOM_PRESETS.length - 1; i >= 0; i--) {
    // SAFETY: i is bounded by ZOOM_PRESETS.length
    if (ZOOM_PRESETS[i]! < currentZoom - 0.01) {
      return ZOOM_PRESETS[i]!;
    }
  }
  // SAFETY: ZOOM_PRESETS is a non-empty constant array
  return ZOOM_PRESETS[0]!;
}

// ============================================================================
// USE WHEEL ZOOM HOOK
// ============================================================================

/**
 * React hook for Ctrl+scroll zoom functionality
 */
export function useWheelZoom(options: UseWheelZoomOptions = {}): UseWheelZoomReturn {
  const {
    initialZoom = DEFAULT_INITIAL_ZOOM,
    minZoom = DEFAULT_MIN_ZOOM,
    maxZoom = DEFAULT_MAX_ZOOM,
    zoomStep = DEFAULT_ZOOM_STEP,
    enabled = true,
    containerRef,
    getCurrentZoom,
    onZoomChange,
    enableKeyboardShortcuts = true,
    preventDefault = true,
  } = options;

  const [zoom, setZoomState] = useState(initialZoom);
  const zoomRef = useRef(zoom);

  // Keep ref in sync with state
  useEffect(() => {
    zoomRef.current = zoom;
  }, [zoom]);

  /**
   * Live current zoom. Controlled mode reads the external owner; uncontrolled
   * mode reads the hook's own mirror.
   */
  const readZoom = useCallback(() => getCurrentZoom?.() ?? zoomRef.current, [getCurrentZoom]);

  /**
   * Set zoom with clamping and callback
   */
  const setZoom = useCallback(
    (newZoom: number) => {
      const clampedZoom = roundZoom(clamp(newZoom, minZoom, maxZoom));
      if (clampedZoom === readZoom()) {
        return;
      }
      // Controlled mode does not own state; the external owner re-renders us.
      // Uncontrolled mode mirrors the new value into the ref synchronously (like
      // DocxEditor's controlled path) so a fast second delta in the same tick
      // composes from the fresh value, not the still-stale pre-render state.
      if (!getCurrentZoom) {
        zoomRef.current = clampedZoom;
        setZoomState(clampedZoom);
      }
      onZoomChange?.(clampedZoom);
    },
    [minZoom, maxZoom, onZoomChange, getCurrentZoom, readZoom],
  );

  /**
   * Zoom in by step
   */
  const zoomIn = useCallback(() => {
    setZoom(readZoom() + zoomStep);
  }, [zoomStep, setZoom, readZoom]);

  /**
   * Zoom out by step
   */
  const zoomOut = useCallback(() => {
    setZoom(readZoom() - zoomStep);
  }, [zoomStep, setZoom, readZoom]);

  /**
   * Reset zoom to initial level
   */
  const resetZoom = useCallback(() => {
    setZoom(initialZoom);
  }, [initialZoom, setZoom]);

  /**
   * Reset zoom to 100%
   */
  const zoomTo100 = useCallback(() => {
    setZoom(1);
  }, [setZoom]);

  /**
   * Zoom to fit width
   */
  const zoomToFit = useCallback(
    (containerWidth: number, contentWidth: number) => {
      if (contentWidth > 0) {
        const fitZoom = containerWidth / contentWidth;
        setZoom(fitZoom);
      }
    },
    [setZoom],
  );

  /**
   * Handle wheel event
   */
  const handleWheel = useCallback(
    (event: WheelEvent) => {
      if (!enabled) {
        return;
      }

      // Check for Ctrl/Cmd key (trackpad pinch also reports ctrlKey)
      const isCtrlOrMeta = event.ctrlKey || event.metaKey;
      if (!isCtrlOrMeta) {
        return;
      }

      // Prevent default browser zoom
      if (preventDefault) {
        event.preventDefault();
      }

      // Scale by a continuous factor proportional to deltaY rather than a fixed
      // step. Trackpad pinch arrives as a stream of small-deltaY ctrlKey wheel
      // events; a discrete step per event would slam between min and max. The
      // exponential keeps zoom multiplicative (symmetric in/out) and smooth for
      // both pinch and discrete mouse-wheel notches. deltaY < 0 = zoom in.
      const delta = event.deltaY;
      if (delta !== 0) {
        setZoom(readZoom() * Math.exp(-delta * WHEEL_ZOOM_SENSITIVITY));
      }
    },
    [enabled, preventDefault, setZoom, readZoom],
  );

  /**
   * Handle keyboard shortcuts
   */
  const handleKeyDown = useCallback(
    (event: KeyboardEvent) => {
      if (!enabled || !enableKeyboardShortcuts) {
        return;
      }

      const isCtrlOrMeta = event.ctrlKey || event.metaKey;
      if (!isCtrlOrMeta) {
        return;
      }

      // Ctrl+0 - Reset zoom to 100%
      if (event.key === "0") {
        event.preventDefault();
        zoomTo100();
        return;
      }

      // Ctrl++ or Ctrl+= - Zoom in
      if (event.key === "+" || event.key === "=") {
        event.preventDefault();
        zoomIn();
        return;
      }

      // Ctrl+- - Zoom out
      if (event.key === "-") {
        event.preventDefault();
        zoomOut();
        return;
      }
    },
    [enabled, enableKeyboardShortcuts, zoomIn, zoomOut, zoomTo100],
  );

  // Attach wheel listener to container. Read `.current` inside the effect (it
  // runs after the DOM is mutated, so the ref is populated by then); handleWheel
  // is rebuilt on render, so the effect re-runs and picks up the container once
  // it mounts on a later render (it starts null until the document loads).
  useEffect(() => {
    const container = containerRef?.current;
    if (!enabled || !container) {
      return;
    }

    // Use passive: false to allow preventDefault
    container.addEventListener("wheel", handleWheel, { passive: false });

    return () => {
      container.removeEventListener("wheel", handleWheel);
    };
  }, [enabled, containerRef, handleWheel]);

  // Attach keyboard listener
  useEffect(() => {
    if (!enabled || !enableKeyboardShortcuts) {
      return;
    }

    document.addEventListener("keydown", handleKeyDown);

    return () => {
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [enabled, enableKeyboardShortcuts, handleKeyDown]);

  const currentZoom = getCurrentZoom ? getCurrentZoom() : zoom;

  return {
    zoom: currentZoom,
    setZoom,
    zoomIn,
    zoomOut,
    resetZoom,
    zoomTo100,
    zoomToFit,
    isMinZoom: currentZoom <= minZoom,
    isMaxZoom: currentZoom >= maxZoom,
    zoomPercent: Math.round(currentZoom * 100),
    handleWheel,
    handleKeyDown,
  };
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get zoom presets
 */
export function getZoomPresets(): number[] {
  return [...ZOOM_PRESETS];
}

/**
 * Find nearest zoom preset
 */
export function findNearestZoomPreset(zoom: number): number {
  return nearestPreset(zoom);
}

/**
 * Get next zoom preset (for zoom in)
 */
export function getNextZoomPreset(zoom: number): number {
  return nextPreset(zoom);
}

/**
 * Get previous zoom preset (for zoom out)
 */
export function getPreviousZoomPreset(zoom: number): number {
  return prevPreset(zoom);
}

/**
 * Format zoom level for display
 */
export function formatZoom(zoom: number): string {
  return `${Math.round(zoom * 100)}%`;
}

/**
 * Parse zoom from percentage string
 */
export function parseZoom(zoomString: string): number | null {
  const match = zoomString.match(/(\d+(\.\d+)?)/);
  if (match) {
    // SAFETY: group 1 always captures in this regex
    const value = Number.parseFloat(match[1]!);
    if (!Number.isNaN(value)) {
      return value / 100;
    }
  }
  return null;
}

/**
 * Check if zoom level is at a preset
 */
export function isZoomPreset(zoom: number): boolean {
  return ZOOM_PRESETS.some((preset) => Math.abs(preset - zoom) < 0.01);
}

/**
 * Clamp zoom to valid range
 */
export function clampZoom(
  zoom: number,
  minZoom: number = DEFAULT_MIN_ZOOM,
  maxZoom: number = DEFAULT_MAX_ZOOM,
): number {
  return roundZoom(clamp(zoom, minZoom, maxZoom));
}
