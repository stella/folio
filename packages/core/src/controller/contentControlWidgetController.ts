/**
 * Framework-neutral lifecycle and state owner for typed content-control pickers.
 *
 * The ProseMirror plugin owns click detection and document transactions. This
 * controller binds that plugin's DOM event to a stable, discriminated snapshot
 * that any adapter can render, then routes a selection back through the same
 * transaction helpers. Framework adapters only subscribe and provide chrome.
 */

import type { EditorView } from "prosemirror-view";

import {
  CONTENT_CONTROL_WIDGET_EVENT_NAME,
  dispatchDatePick,
  dispatchDropdownPick,
} from "../prosemirror/plugins/contentControlWidgets";
import type { ContentControlWidgetEvent } from "../prosemirror/plugins/contentControlWidgets";
import { Subscribable } from "../managers/Subscribable";

export type ContentControlListItem = {
  displayText: string;
  value: string;
};

export type ContentControlPickerPosition = {
  x: number;
  y: number;
};

export type ContentControlWidgetSnapshot =
  | { status: "closed" }
  | {
      status: "dropdown";
      items: readonly ContentControlListItem[];
      pmPos: number;
      position: ContentControlPickerPosition;
    }
  | {
      status: "date";
      currentValue: string | null;
      pmPos: number;
      position: ContentControlPickerPosition;
    };

export type ContentControlWidgetControllerOptions = {
  onRefused?:
    | ((event: Extract<ContentControlWidgetEvent, { kind: "refused" }>) => void)
    | undefined;
};

const CLOSED_SNAPSHOT: ContentControlWidgetSnapshot = { status: "closed" };
const PICKER_GAP_PX = 4;

const isListItem = (value: unknown): value is ContentControlListItem => {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  if (!("displayText" in value) || !("value" in value)) {
    return false;
  }
  return typeof value.displayText === "string" && typeof value.value === "string";
};

export const parseContentControlListItems = (raw: string | undefined): ContentControlListItem[] => {
  if (!raw) {
    return [];
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return [];
  }
  if (!Array.isArray(parsed)) {
    return [];
  }
  return parsed.filter(isListItem);
};

const pickerPosition = (
  event: Extract<ContentControlWidgetEvent, { kind: "dropdownOpen" | "datePick" }>,
): ContentControlPickerPosition => {
  const rect = event.anchor.getBoundingClientRect();
  return { x: rect.left, y: rect.bottom + PICKER_GAP_PX };
};

const isContentControlWidgetEvent = (value: unknown): value is ContentControlWidgetEvent => {
  if (typeof value !== "object" || value === null || !("kind" in value)) {
    return false;
  }
  if (
    !("tag" in value) ||
    typeof value.tag !== "string" ||
    !("pmPos" in value) ||
    typeof value.pmPos !== "number" ||
    !("anchor" in value) ||
    typeof value.anchor !== "object" ||
    value.anchor === null ||
    !("getBoundingClientRect" in value.anchor) ||
    typeof value.anchor.getBoundingClientRect !== "function"
  ) {
    return false;
  }
  if (value.kind === "dropdownOpen") {
    return (
      "sdtType" in value &&
      (value.sdtType === "dropdown" || value.sdtType === "comboBox") &&
      "listItemsJson" in value &&
      (value.listItemsJson === undefined || typeof value.listItemsJson === "string")
    );
  }
  if (value.kind === "datePick") {
    return (
      "currentValue" in value &&
      (value.currentValue === undefined || typeof value.currentValue === "string")
    );
  }
  return value.kind === "refused" && "sdtType" in value && "error" in value;
};

export class ContentControlWidgetController extends Subscribable<ContentControlWidgetSnapshot> {
  private readonly options: ContentControlWidgetControllerOptions;
  private view: EditorView | null = null;
  private viewDom: HTMLElement | null = null;
  private ownerWindow: Window | null = null;

  constructor(options: ContentControlWidgetControllerOptions = {}) {
    super(CLOSED_SNAPSHOT);
    this.options = options;
  }

  bind(view: EditorView | null): void {
    if (this.view === view) {
      return;
    }
    this.unbindView();
    this.view = view;
    this.viewDom = view?.dom ?? null;
    this.viewDom?.addEventListener(CONTENT_CONTROL_WIDGET_EVENT_NAME, this.onDomEvent);
  }

  handleWidgetEvent(event: ContentControlWidgetEvent): void {
    if (event.kind === "refused") {
      this.close();
      this.options.onRefused?.(event);
      return;
    }
    if (event.kind === "dropdownOpen") {
      this.open({
        status: "dropdown",
        items: parseContentControlListItems(event.listItemsJson),
        pmPos: event.pmPos,
        position: pickerPosition(event),
      });
      return;
    }
    this.open({
      status: "date",
      currentValue: event.currentValue ?? null,
      pmPos: event.pmPos,
      position: pickerPosition(event),
    });
  }

  pickDropdown(value: string): boolean {
    const snapshot = this.getSnapshot();
    if (snapshot.status !== "dropdown" || !this.view) {
      return false;
    }
    const changed = dispatchDropdownPick(this.view, snapshot.pmPos, value);
    this.close();
    return changed;
  }

  pickDate(value: string): boolean {
    const snapshot = this.getSnapshot();
    if (snapshot.status !== "date" || !this.view || value.length === 0) {
      return false;
    }
    const changed = dispatchDatePick(this.view, snapshot.pmPos, value);
    this.close();
    return changed;
  }

  close(): void {
    this.removeDismissListeners();
    if (this.getSnapshot().status !== "closed") {
      this.setSnapshot(CLOSED_SNAPSHOT);
    }
  }

  destroy(): void {
    this.unbindView();
    this.view = null;
  }

  private readonly onDomEvent = (event: Event): void => {
    if (!("detail" in event)) {
      return;
    }
    const detail: unknown = event.detail;
    if (isContentControlWidgetEvent(detail)) {
      this.handleWidgetEvent(detail);
    }
  };

  private readonly onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      this.close();
    }
  };

  private readonly onScroll = (): void => this.close();

  private open(snapshot: Exclude<ContentControlWidgetSnapshot, { status: "closed" }>): void {
    this.removeDismissListeners();
    this.ownerWindow = this.viewDom?.ownerDocument.defaultView ?? null;
    this.ownerWindow?.addEventListener("keydown", this.onKeyDown);
    this.ownerWindow?.addEventListener("scroll", this.onScroll, true);
    this.setSnapshot(snapshot);
  }

  private removeDismissListeners(): void {
    this.ownerWindow?.removeEventListener("keydown", this.onKeyDown);
    this.ownerWindow?.removeEventListener("scroll", this.onScroll, true);
    this.ownerWindow = null;
  }

  private unbindView(): void {
    this.close();
    this.viewDom?.removeEventListener(CONTENT_CONTROL_WIDGET_EVENT_NAME, this.onDomEvent);
    this.viewDom = null;
  }
}
