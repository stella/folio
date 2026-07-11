/**
 * Wrap a React event handler so it only fires when the event target
 * is a DOM descendant of the element handling the event.
 *
 * React forwards synthetic events through the parent React tree even
 * when descendants are rendered via createPortal. That makes it unsafe
 * for a parent element to attach `onMouseDown`, `onClick`, `onFocus`,
 * etc. that side-effect (focus steal, preventDefault, selection changes)
 * under the assumption "the target lives inside me" — a portaled popup
 * (Dialog, Combobox, Tooltip) silently bypasses the DOM boundary and
 * triggers the side effect, dismissing itself.
 *
 * Only use for event types whose `target` is the meaningful subject of
 * the event: pointer/mouse/touch/click and `focus`. Do **not** wrap
 * `onBlur`: blur's `target` is the element losing focus, not the new
 * focus destination, so the containment test cannot answer "is focus
 * leaving for a portaled child?". For that case test `relatedTarget`
 * directly inside the handler.
 *
 * The `require-contained-handler` oxlint rule enforces this on any JSX
 * element carrying both `ref={…}` and one of the watched handler props.
 *
 * @example
 *   return (
 *     <div
 *       onMouseDown={containedHandler((e) => {
 *         e.preventDefault();
 *         focusEditor();
 *       })}
 *     >
 *       ...
 *     </div>
 *   );
 */
export const containedHandler =
  <E extends { currentTarget: EventTarget & Element; target: unknown }>(
    handler: ((event: E) => void) | undefined,
  ) =>
  (event: E): void => {
    if (handler === undefined) {
      return;
    }
    if (event.target instanceof Node && !event.currentTarget.contains(event.target)) {
      return;
    }
    handler(event);
  };
