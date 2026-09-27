/**
 * The shared behaviour of a panel opened as a drawer over the page: a scrim
 * that closes it on an outside press, Escape that closes it, and focus that
 * moves into the drawer when it opens and back to where it came from when it
 * closes.
 */

import { type RefObject, useEffect, useRef } from "react";

const FOCUSABLE =
  'button:not([disabled]), [href], input:not([disabled]), textarea:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])';

/**
 * What a drawer focuses when it opens: its current item (else the list's tab
 * stop, else its first control), or the drawer itself, for one whose controls
 * are scattered down the document.
 */
export type DrawerInitialFocus = "item" | "container";

const LIST_TAB_STOP = ':is(ol, ul) :is(button, [href], [tabindex]):not([tabindex="-1"])';

const initialFocusTarget = (container: HTMLElement, initial: DrawerInitialFocus): HTMLElement => {
  if (initial === "container") {
    return container;
  }
  return (
    container.querySelector<HTMLElement>('[aria-current="true"]') ??
    container.querySelector<HTMLElement>(LIST_TAB_STOP) ??
    container.querySelector<HTMLElement>(FOCUSABLE) ??
    container
  );
};

/**
 * While `onClose` is set, `container` is an open drawer: focus moves into it,
 * Escape calls `onClose`, and once it closes focus returns to what had it
 * before, if that is still in the document.
 */
export const useDrawerFocus = (
  container: RefObject<HTMLElement | null>,
  onClose: (() => void) | null,
  initial: DrawerInitialFocus = "item",
) => {
  const onCloseRef = useRef(onClose);
  useEffect(() => {
    onCloseRef.current = onClose;
  }, [onClose]);
  const active = onClose !== null;

  useEffect(() => {
    if (!active) {
      return undefined;
    }
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const element = container.current;
    if (element) {
      initialFocusTarget(element, initial).focus({ preventScroll: true });
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented) {
        return;
      }
      event.preventDefault();
      onCloseRef.current?.();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("keydown", handleKeyDown);
      if (previous?.isConnected) {
        previous.focus({ preventScroll: true });
      }
    };
  }, [active, container, initial]);
};

type PanelScrimProps = {
  onDismiss: () => void;
};

/** Dims the page under an open drawer; pressing it closes the drawer. */
export const PanelScrim = ({ onDismiss }: PanelScrimProps) => {
  // Escape closes the drawer for keyboard users; the scrim is the pointer's
  // equivalent, hidden from assistive technology like any backdrop.
  return (
    // oxlint-disable-next-line jsx-a11y/no-static-element-interactions -- backdrop; Escape is the keyboard path
    <div
      aria-hidden="true"
      className="folio-panel-scrim"
      data-testid="folio-panel-scrim"
      onMouseDown={(event) => {
        event.preventDefault();
        onDismiss();
      }}
    />
  );
};
