/**
 * Collapse a burst of calls into one, `delayMs` after the last. A folio write
 * replaces the file by rename, and an editor saves in several steps, so one
 * change arrives as several watcher events.
 */
export type Debounced = {
  readonly schedule: () => void;
  readonly dispose: () => void;
};

export const debounce = (run: () => void, delayMs: number): Debounced => {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return {
    schedule: () => {
      if (timer !== undefined) clearTimeout(timer);
      timer = setTimeout(() => {
        timer = undefined;
        run();
      }, delayMs);
    },
    dispose: () => {
      if (timer !== undefined) clearTimeout(timer);
      timer = undefined;
    },
  };
};
