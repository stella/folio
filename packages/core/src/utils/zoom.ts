/**
 * Canonical zoom limits shared by every folio adapter (React, Vue, ...).
 *
 * Adapters must source their zoom range from here rather than hard-coding their
 * own, so the reachable zoom span stays identical across framework hosts.
 * Historically React clamped 0.25-4x while Vue clamped 0.5-2x; this module is
 * the single source of truth that keeps them from drifting again.
 *
 * Note: the discrete zoom-level menu each adapter shows in its toolbar dropdown
 * is a separate, deliberately curated subset (50-200%); it is not derived from
 * these limits.
 */

// ============================================================================
// CONSTANTS
// ============================================================================

/** Minimum reachable zoom level (25%). */
export const ZOOM_MIN = 0.25;

/** Maximum reachable zoom level (400%). */
export const ZOOM_MAX = 4;

/** Step applied per zoom-in / zoom-out action and per keyboard shortcut. */
export const ZOOM_STEP = 0.1;
