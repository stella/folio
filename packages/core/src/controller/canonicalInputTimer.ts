import type { EditorView } from "prosemirror-view";

/** Shared by timer producers and browser diagnostics, regardless of bundling. */
export const CANONICAL_INPUT_TIMER_OWNER = "folio.canonical-input-timer";

/** Delay after native composition end before the canonical session settles it. */
export const CANONICAL_COMPOSITION_FLUSH_DELAY_MS = 25;

/** Mark the scheduled wrapper, never the caller's potentially shared callback. */
export const setCanonicalInputTimer = (callback: () => void, delay: number) => {
  const ownedCallback = () => callback();
  Object.defineProperty(ownedCallback, Symbol.for(CANONICAL_INPUT_TIMER_OWNER), { value: true });
  return setTimeout(ownedCallback, delay);
};

/**
 * Run `retry` once the view's canonical composition has settled. The session
 * settles on an equal-delay timer armed at native composition end, so a retry
 * armed at or after that point runs after it. Returns a disposer.
 */
export const afterCanonicalCompositionSettles = (
  view: EditorView,
  retry: () => void,
): (() => void) => {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const schedule = () => {
    timer = setCanonicalInputTimer(retry, CANONICAL_COMPOSITION_FLUSH_DELAY_MS);
  };
  const onEnd = () => {
    view.dom.removeEventListener("compositionend", onEnd);
    schedule();
  };
  if (view.composing) view.dom.addEventListener("compositionend", onEnd);
  else schedule();
  return () => {
    view.dom.removeEventListener("compositionend", onEnd);
    if (timer !== undefined) clearTimeout(timer);
  };
};
