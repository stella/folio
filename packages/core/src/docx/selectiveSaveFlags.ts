/**
 * Selective Save Feature Flags
 *
 * Operational controls for the selective save path. Selective save is on by
 * default: every save attempts it first and falls back to a full repack
 * whenever the patch-safety checks refuse (structural changes, untracked
 * changes, oversized buffers, or an unmodelled part). Hosts that need the
 * previous behavior can opt back into full repacking on every save.
 *
 * - `selectiveSave` — gate the selective branch entirely.
 * - `selectiveSaveTripwire` — run selective + full and compare bytes for CI
 *   observability. Never blocks the user-visible save.
 * - `selectiveSaveMaxBytes` — refuse the selective path for original buffers
 *   above this size, since holding them in memory plus the JSZip overhead is
 *   the dominant cost.
 */

export type FolioSelectiveSaveFlags = {
  /**
   * Enable the selective save path. Default: true. A save always succeeds:
   * when the patch-safety checks refuse, the caller falls back to a full
   * repack automatically. Set to `false` to always full-repack.
   */
  selectiveSave?: boolean;
  /**
   * Run selective save and full repack on every save, compare bytes, and emit
   * a `TripwireResult` via `onSelectiveSaveTripwire`. Independent of
   * `selectiveSave`: the tripwire always observes both paths.
   */
  selectiveSaveTripwire?: boolean;
  /** Maximum original buffer size, in bytes, for which selective save is allowed. */
  selectiveSaveMaxBytes?: number;
};

export const DEFAULT_SELECTIVE_SAVE_MAX_BYTES = 100 * 1024 * 1024;

export type ResolvedSelectiveSaveFlags = {
  selectiveSave: boolean;
  selectiveSaveTripwire: boolean;
  selectiveSaveMaxBytes: number;
};

export function resolveSelectiveSaveFlags(
  flags: FolioSelectiveSaveFlags | undefined,
): ResolvedSelectiveSaveFlags {
  return {
    selectiveSave: flags?.selectiveSave ?? true,
    selectiveSaveTripwire: flags?.selectiveSaveTripwire ?? false,
    selectiveSaveMaxBytes: flags?.selectiveSaveMaxBytes ?? DEFAULT_SELECTIVE_SAVE_MAX_BYTES,
  };
}
