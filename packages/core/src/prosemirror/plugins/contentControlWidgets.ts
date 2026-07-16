/**
 * Interactive content-control widget delegation.
 *
 * Painted blocks inside a typed SDT carry `data-sdt-type` (checkbox /
 * dropdown / date) and `data-sdt-tag` thanks to `applySdtDataAttrs`. This
 * plugin watches editor-region clicks: when the user clicks inside a typed
 * control it dispatches the appropriate `setContentControlValueTr` so the
 * mutation goes through the normal undo stack.
 *
 * The checkbox path is fully wired (a click toggles the modeled state and
 * the rendered glyph). Dropdown and date paths emit the dispatch hook so the
 * higher-level UI (a popover menu / date picker rendered by the editor
 * shell) can subscribe; the plugin does not own the menu/picker chrome.
 */

import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { ContentControlLockedError, ContentControlTypeError } from "../../content-controls";
import { findBlockSdtMatch, setContentControlValueTr } from "../commands/contentControls";

export type ContentControlWidgetAnchor = {
  getBoundingClientRect: () => { bottom: number; left: number };
};

export type ContentControlWidgetEvent =
  | {
      kind: "dropdownOpen";
      tag: string;
      /** PM position of the clicked SDT instance — disambiguates duplicates. */
      pmPos: number;
      sdtType: "dropdown" | "comboBox";
      anchor: ContentControlWidgetAnchor;
      listItemsJson: string | undefined;
    }
  | {
      kind: "datePick";
      tag: string;
      pmPos: number;
      anchor: ContentControlWidgetAnchor;
      currentValue: string | undefined;
    }
  | {
      /**
       * The plugin refused a click because the control's `w:lock` or type
       * forbade the interaction. The shell decides how to surface this —
       * toast, telemetry, or no-op. The plugin itself never logs.
       */
      kind: "refused";
      tag: string;
      pmPos: number;
      sdtType: string;
      anchor: ContentControlWidgetAnchor;
      error: ContentControlLockedError | ContentControlTypeError;
    };

export type ContentControlWidgetCallback = (event: ContentControlWidgetEvent) => void;

type ContentControlWidgetClickEvent = {
  preventDefault: () => void;
  target: unknown;
};

type ContentControlWidgetView = Pick<EditorView, "dispatch" | "state">;

type HandleContentControlWidgetClickOptions = {
  event: ContentControlWidgetClickEvent;
  onEvent: ContentControlWidgetCallback;
  view: ContentControlWidgetView;
};

type ContentControlWidgetElement = ContentControlWidgetAnchor & {
  dataset: DOMStringMap;
};

export const contentControlWidgetsPluginKey = new PluginKey<unknown>("contentControlWidgets");

/**
 * CustomEvent name dispatched on the editor view DOM whenever the plugin
 * needs to surface a `ContentControlWidgetEvent` to an adapter shell. The
 * adapter subscribes via `addEventListener` and renders the matching chrome.
 *
 * Using a CustomEvent keeps the plugin framework-neutral and makes the
 * adapter's responsibility explicit: receive event → render UI →
 * call dispatchDropdownPick / dispatchDatePick.
 */
export const CONTENT_CONTROL_WIDGET_EVENT_NAME = "folio:content-control-widget";

const findClosest = (target: unknown, selector: string): unknown => {
  if (
    typeof target !== "object" ||
    target === null ||
    !("closest" in target) ||
    typeof target.closest !== "function"
  ) {
    return null;
  }
  return target.closest(selector);
};

const isContentControlWidgetElement = (value: unknown): value is ContentControlWidgetElement => {
  if (
    typeof value !== "object" ||
    value === null ||
    !("dataset" in value) ||
    typeof value.dataset !== "object" ||
    value.dataset === null ||
    !("getBoundingClientRect" in value) ||
    typeof value.getBoundingClientRect !== "function"
  ) {
    return false;
  }
  return true;
};

const findSdtAncestor = (target: unknown): ContentControlWidgetElement | null => {
  const candidate = findClosest(target, "[data-sdt-type]");
  return isContentControlWidgetElement(candidate) ? candidate : null;
};

/**
 * Interpret one click from either the live ProseMirror DOM or the painted page
 * tree. Both surfaces carry the same `data-sdt-*` projection, so adapters only
 * need to forward the click and render the resulting typed event.
 */
export const handleContentControlWidgetClick = ({
  event,
  onEvent,
  view,
}: HandleContentControlWidgetClickOptions): boolean => {
  const anchor = findSdtAncestor(event.target);
  if (!anchor) {
    return false;
  }
  const tag = anchor.dataset["sdtTag"] ?? "";
  const sdtType = anchor.dataset["sdtType"];
  // pmPos is what addresses the clicked instance unambiguously. The
  // painter stamps it from the SdtGroup; tag is kept on the event
  // for telemetry but is no longer the addressing key.
  const pmPosRaw = anchor.dataset["sdtPmPos"];
  const pmPos = pmPosRaw ? Number.parseInt(pmPosRaw, 10) : Number.NaN;
  if (!sdtType || Number.isNaN(pmPos)) {
    return false;
  }
  // The lock check happens inside the transaction helper; we just
  // surface the click intent. Errors thrown by the helper are
  // caught here so the editor stays interactive even on refusal.
  try {
    if (sdtType === "checkbox") {
      const current = anchor.dataset["sdtChecked"] === "true";
      const tr = setContentControlValueTr(
        view.state,
        { pmPos },
        {
          kind: "checkbox",
          checked: !current,
        },
      );
      if (tr) {
        view.dispatch(tr);
        event.preventDefault();
        return true;
      }
    } else if (sdtType === "dropdown" || sdtType === "comboBox") {
      const refusal = lockRefusalFor(view, pmPos);
      if (refusal) {
        onEvent({
          kind: "refused",
          tag,
          pmPos,
          sdtType,
          anchor,
          error: refusal,
        });
        event.preventDefault();
        return true;
      }
      onEvent({
        kind: "dropdownOpen",
        tag,
        pmPos,
        sdtType,
        anchor,
        listItemsJson: anchor.dataset["sdtListItems"],
      });
      event.preventDefault();
      return true;
    } else if (sdtType === "date") {
      const refusal = lockRefusalFor(view, pmPos);
      if (refusal) {
        onEvent({
          kind: "refused",
          tag,
          pmPos,
          sdtType,
          anchor,
          error: refusal,
        });
        event.preventDefault();
        return true;
      }
      onEvent({
        kind: "datePick",
        tag,
        pmPos,
        anchor,
        currentValue: findBlockSdtMatch(view.state.doc, { pmPos })?.node.textContent,
      });
      event.preventDefault();
      return true;
    }
  } catch (error) {
    if (error instanceof ContentControlLockedError || error instanceof ContentControlTypeError) {
      // Refusal is expected: emit it through the typed callback and let
      // the editor shell decide how to surface it.
      onEvent({
        kind: "refused",
        tag,
        pmPos,
        sdtType,
        anchor,
        error,
      });
      return true;
    }
    throw error;
  }
  return false;
};

export function createContentControlWidgetsPlugin(
  onEvent: ContentControlWidgetCallback = () => undefined,
): Plugin {
  const emit = (view: { dom: HTMLElement }, payload: ContentControlWidgetEvent): void => {
    onEvent(payload);
    view.dom.dispatchEvent(
      new CustomEvent(CONTENT_CONTROL_WIDGET_EVENT_NAME, {
        detail: payload,
        bubbles: true,
      }),
    );
  };
  return new Plugin({
    key: contentControlWidgetsPluginKey,
    props: {
      handleDOMEvents: {
        click(view, event) {
          return handleContentControlWidgetClick({
            view,
            event,
            onEvent: (payload) => emit(view, payload),
          });
        },
      },
    },
    view() {
      // Resolves the helper at plugin-find time so the picker shell can
      // dispatch a tx by tag without re-implementing the lock/type rules.
      return {};
    },
  });
}

/**
 * Stable resolver helpers exposed for the UI shell so a popover menu or
 * date picker can dispatch the same transaction the plugin would.
 *
 * Both take the PM `pmPos` of the clicked SDT (from the widget event's
 * `pmPos` field) so the dispatch lands on the exact instance the user
 * interacted with — duplicate-tag SDTs are no longer ambiguous.
 */
export function dispatchDropdownPick(view: EditorView, pmPos: number, value: string): boolean {
  const match = findBlockSdtMatch(view.state.doc, { pmPos });
  if (!match) {
    return false;
  }
  try {
    const tr = setContentControlValueTr(
      view.state,
      { pmPos },
      {
        kind: "dropdown",
        value,
      },
    );
    if (!tr) {
      return false;
    }
    view.dispatch(tr);
    return true;
  } catch (error) {
    // The click-time preflight should have caught lock refusals, but the
    // doc could have changed mid-picker. Swallow the typed refusal so the
    // adapter UI does not crash; return false so the caller knows the
    // change was rejected.
    if (error instanceof ContentControlLockedError || error instanceof ContentControlTypeError) {
      return false;
    }
    throw error;
  }
}

export function dispatchDatePick(view: EditorView, pmPos: number, date: string): boolean {
  try {
    const tr = setContentControlValueTr(
      view.state,
      { pmPos },
      {
        kind: "date",
        date,
      },
    );
    if (!tr) {
      return false;
    }
    view.dispatch(tr);
    return true;
  } catch (error) {
    if (error instanceof ContentControlLockedError || error instanceof ContentControlTypeError) {
      return false;
    }
    throw error;
  }
}

/**
 * Look up the clicked SDT by pmPos and return a typed refusal when the
 * lock state forbids content mutation (`contentLocked`,
 * `sdtContentLocked`). Returning `null` means "go ahead and open the
 * picker." Preflight at click time so locked dropdowns / date pickers
 * don't flash open just to close again when the user makes a selection.
 */
function lockRefusalFor(
  view: Pick<EditorView, "state">,
  pmPos: number,
): ContentControlLockedError | null {
  const match = findBlockSdtMatch(view.state.doc, { pmPos });
  if (!match) {
    return null;
  }
  const lock = match.node.attrs["lock"];
  if (lock !== "contentLocked" && lock !== "sdtContentLocked") {
    return null;
  }
  const tag = match.node.attrs["tag"];
  const alias = match.node.attrs["alias"];
  return new ContentControlLockedError({
    message: `Control "${tag ?? alias ?? "(unnamed)"}" has w:lock=${lock}.`,
    lock,
    ...(typeof tag === "string" ? { tag } : {}),
    ...(typeof alias === "string" ? { alias } : {}),
  });
}
