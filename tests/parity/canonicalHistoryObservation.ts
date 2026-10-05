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

type CanonicalNativeEventContext = {
  propagation: "capture" | "bubble";
  trusted: boolean;
  defaultPrevented: boolean;
  target: "editor" | "other";
  state: CanonicalHistoryState;
};

export const CANONICAL_NATIVE_EVENT_TYPES = [
  "compositionstart",
  "compositionupdate",
  "compositionend",
  "beforeinput",
  "input",
] as const;
type NativeEventType = (typeof CANONICAL_NATIVE_EVENT_TYPES)[number];

export type CanonicalNativeEvent = CanonicalNativeEventContext &
  (
    | { type: Extract<NativeEventType, `composition${string}`>; data: string }
    | {
        type: Exclude<NativeEventType, `composition${string}`>;
        data: string | null;
        inputType: string;
        composing: boolean;
      }
  );

export type CanonicalImeDriverCall = {
  operation: "start" | "update" | "commit" | "cancel";
  text: string;
  before: CanonicalHistoryState | null;
  after: CanonicalHistoryState | null;
};

export type CanonicalHistoryObservation = {
  before: CanonicalHistoryState;
  after?: CanonicalHistoryState;
  keys: CanonicalHistoryKeyEvent[];
  nativeEvents: CanonicalNativeEvent[];
  driverCalls: CanonicalImeDriverCall[];
  capture: { status: "complete" } | { status: "unavailable"; message: string };
};

declare global {
  var __folioCanonicalHistoryKeys: CanonicalHistoryKeyEvent[] | undefined;
  var __folioCanonicalNativeEvents: CanonicalNativeEvent[] | undefined;
  var __folioCanonicalImeDriverCalls: CanonicalImeDriverCall[] | undefined;
}
