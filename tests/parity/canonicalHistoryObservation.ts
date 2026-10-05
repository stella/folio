export type CanonicalHistoryState = {
  version: number | null;
  focused: boolean;
  composing: boolean;
  canUndo: boolean;
  canRedo: boolean;
  selection: { anchor: number; head: number } | null;
  canonicalSelection: { anchor: number; head: number } | null;
};

export type CanonicalHistoryKeyEvent = {
  propagation: "capture" | "bubble";
  key: string;
  control: boolean;
  meta: boolean;
  shift: boolean;
  defaultPrevented: boolean;
  target: "editor" | "other";
  state: CanonicalHistoryState;
};

export type CanonicalHistoryObservation = {
  before: CanonicalHistoryState;
  after?: CanonicalHistoryState;
  keys: CanonicalHistoryKeyEvent[];
  capture: { status: "complete" } | { status: "unavailable"; message: string };
};

declare global {
  var __folioCanonicalHistoryKeys: CanonicalHistoryKeyEvent[] | undefined;
}
