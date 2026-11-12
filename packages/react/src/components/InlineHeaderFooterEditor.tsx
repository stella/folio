/**
 * InlineHeaderFooterEditor — UI chrome for HF editing.
 *
 * Before the HF editing unification, this component mounted its own visible
 * ProseMirror EditorView positioned over the painted HF area. The painter
 * underneath was hidden via `visibility: hidden` and the visible PM owned
 * editing. That model produced edit/non-edit geometry drift, field-insert
 * extra keystrokes, scroll-to-page-1 on cross-page HF edits, and imprecise
 * caret placement.
 *
 * Post-unification (eigenpal#611), the painter is the sole visible HF
 * renderer in both edit and non-edit modes; user input flows through the
 * persistent off-screen EditorView mounted by `HiddenHeaderFooterPMs`.
 * This component is now just chrome: a positioned label, an options menu
 * (Insert PAGE / NUMPAGES / Remove / Close), and the separator bar.
 *
 * The ref still exposes `getView` / `focus` / `undo` / `redo` so existing
 * callers (`useHeaderFooterEditor`, toolbar wiring) keep working — those
 * methods now delegate to the active HF PM via `getActiveView`.
 */

import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import type {
  CSSProperties,
  ForwardRefExoticComponent,
  RefAttributes,
} from "react";

import { redo, undo } from "prosemirror-history";
import type { EditorView } from "prosemirror-view";
import { useTranslations } from "use-intl";

import { PAINTER_PAINTED_EVENT } from "@stll/folio-core/layout-painter/renderPage";
import { schema } from "@stll/folio-core/prosemirror";
import "prosemirror-view/style/prosemirror.css";

import { useEqualState } from "../hooks/keepEqual";

// ============================================================================
// TYPES
// ============================================================================

export type InlineHeaderFooterEditorProps = {
  /** Whether editing header or footer */
  position: "header" | "footer";
  /**
   * The selector of the painted element to overlay under `parentElement`
   * (`.layout-page-header` / `.layout-page-footer`). The painter replaces
   * that element when it repaints the page, so the chrome looks it up again
   * on every paint.
   */
  targetSelector: string;
  /** The positioning parent element (the div wrapping PagedEditor) */
  parentElement: HTMLElement;
  /**
   * Returns the persistent hidden HF EditorView the user is currently
   * editing. The chrome's ref methods (`getView`, `focus`, `undo`, `redo`)
   * and the options menu's field inserts all route through this.
   */
  getActiveView: () => EditorView | null;
  /** Callback when editing is cancelled (Escape / Close button) */
  onClose: () => void;
  /** Callback to remove the header/footer entirely */
  onRemove?: () => void;
};

export type InlineHeaderFooterEditorRef = {
  /** Get the active HF PM EditorView (the persistent hidden one). */
  getView: () => EditorView | null;
  /** Focus the active HF PM. */
  focus: () => void;
  /** Undo on the active HF PM. */
  undo: () => boolean;
  /** Redo on the active HF PM. */
  redo: () => boolean;
};

// ============================================================================
// STYLES
// ============================================================================

const separatorBarStyle: CSSProperties = {
  position: "absolute",
  left: 0,
  right: 0,
  // Anchor the bar's bottom edge to the container's top so the "Header" label
  // and Options sit ABOVE the painted HF content. The painter shows through the
  // overlay now, so an in-flow bar would render right on top of the first row
  // of HF text; hoisting it outside the content box keeps them from overlapping.
  bottom: "100%",
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  padding: "2px 0",
  fontSize: 11,
  color: "var(--doc-link)",
  userSelect: "none",
  // Container is pointer-events: none so the painter beneath receives clicks;
  // restore events on the chrome bar so the label + Options stay interactive.
  pointerEvents: "auto",
};

// Footer bar mirrors the header bar but sits BELOW its painted content (its top
// edge anchored to the container's bottom) so it never overlaps the footer text.
const footerBarStyle: CSSProperties = {
  ...separatorBarStyle,
  bottom: "auto",
  top: "100%",
};

const labelStyle: CSSProperties = {
  fontWeight: 500,
  letterSpacing: 0.3,
};

// ============================================================================
// COMPONENT
// ============================================================================

/* eslint-disable prefer-arrow-callback -- preserve the component name in React DevTools without reindenting the implementation. */
const InlineHeaderFooterEditorComponent = forwardRef<
  InlineHeaderFooterEditorRef,
  InlineHeaderFooterEditorProps
>(function InlineHeaderFooterEditor(
  {
    position,
    targetSelector,
    parentElement,
    getActiveView,
    onClose,
    onRemove,
  }: InlineHeaderFooterEditorProps,
  ref,
) {
  const t = useTranslations("folio");
  const [showOptions, setShowOptions] = useState(false);
  const optionsRef = useRef<HTMLDivElement>(null);

  // A repaint writes the same position on most keys; an equal one must not
  // commit.
  const [overlayPos, setOverlayPos] = useEqualState<{
    top: number;
    left: number;
    width: number;
    height: number;
  } | null>(null);

  useEffect(() => {
    let observed: HTMLElement | null = null;
    // The HF box grows/shrinks as the user types, and the footer bar is
    // anchored to the box bottom (`top: 100%`), so the overlay height must
    // track the box — recompute on target resize, not just scroll/window.
    const resizeObserver = new ResizeObserver(() => {
      computePosition();
    });
    // The painter replaces the painted slot when it repaints its page, so the
    // position is read from the slot that is painted now, never one held
    // from an earlier paint: a replaced one is detached and measures zero.
    const computePosition = () => {
      const target = parentElement.querySelector<HTMLElement>(targetSelector);
      if (target !== observed) {
        if (observed !== null) resizeObserver.unobserve(observed);
        if (target !== null) resizeObserver.observe(target);
        observed = target;
      }
      if (target === null) return;
      const parentRect = parentElement.getBoundingClientRect();
      const targetRect = target.getBoundingClientRect();
      setOverlayPos({
        top: targetRect.top - parentRect.top + parentElement.scrollTop,
        left: targetRect.left - parentRect.left + parentElement.scrollLeft,
        width: targetRect.width,
        height: targetRect.height,
      });
    };
    computePosition();

    // The painted event bubbles from the pages container to this parent.
    parentElement.addEventListener(PAINTER_PAINTED_EVENT, computePosition);
    const scrollParent =
      parentElement.closest('[style*="overflow"]') || parentElement;
    scrollParent.addEventListener("scroll", computePosition);
    window.addEventListener("resize", computePosition);
    return () => {
      resizeObserver.disconnect();
      parentElement.removeEventListener(PAINTER_PAINTED_EVENT, computePosition);
      scrollParent.removeEventListener("scroll", computePosition);
      window.removeEventListener("resize", computePosition);
    };
  }, [parentElement, targetSelector, setOverlayPos]);

  // Latest getActiveView, read through a ref so the focus loop below is never
  // stuck on a stale mount-time closure if the active HF view resolves later.
  const getActiveViewRef = useRef(getActiveView);
  useLayoutEffect(() => {
    getActiveViewRef.current = getActiveView;
  }, [getActiveView]);

  // Move focus into the persistent HF view when the chrome mounts so typing
  // starts immediately after double-click. A single focus() is unreliable: the
  // HF view may not exist yet at mount (HiddenHeaderFooterPMs creates views in
  // its own effect), AND entering edit mode triggers a relayout/repaint that
  // re-grabs focus for the body PM — so a one-shot focus() gets stolen back and
  // typing lands in the body. Keep nudging focus into the HF view across a short
  // rAF window until it actually holds focus (or the window ends).
  useEffect(() => {
    const MAX_FOCUS_ATTEMPTS = 30;
    let attempts = 0;
    let rafId: number | null = null;
    const tryFocus = () => {
      const view = getActiveViewRef.current();
      if (view && !view.hasFocus()) {
        view.focus();
      }
      attempts++;
      // Keep watching for the whole window, not just until the first successful
      // focus: entering edit mode triggers a relayout/repaint that can clear or
      // steal focus *after* it first lands, so re-focus on any frame where the
      // HF view has lost it.
      if (attempts < MAX_FOCUS_ATTEMPTS) {
        rafId = requestAnimationFrame(tryFocus);
      }
    };
    rafId = requestAnimationFrame(tryFocus);
    return () => {
      if (rafId !== null) {
        cancelAnimationFrame(rafId);
      }
    };
  }, []);

  const handleClose = useCallback(() => {
    onClose();
  }, [onClose]);

  useEffect(() => {
    function handleKeyDown(e: KeyboardEvent) {
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        handleClose();
      }
    }
    document.addEventListener("keydown", handleKeyDown, true);
    return () => document.removeEventListener("keydown", handleKeyDown, true);
  }, [handleClose]);

  useEffect(() => {
    if (!showOptions) {
      return;
    }
    function handleClick(e: MouseEvent) {
      if (
        optionsRef.current &&
        !optionsRef.current.contains(e.target as Node)
      ) {
        setShowOptions(false);
      }
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, [showOptions]);

  useImperativeHandle(ref, () => ({
    getView: () => getActiveView(),
    focus: () => getActiveView()?.focus(),
    undo: () => {
      const view = getActiveView();
      if (!view) {
        return false;
      }
      return undo(view.state, view.dispatch);
    },
    redo: () => {
      const view = getActiveView();
      if (!view) {
        return false;
      }
      return redo(view.state, view.dispatch);
    },
  }));

  const label =
    position === "header" ? t("headerFooter.header") : t("headerFooter.footer");

  if (!overlayPos) {
    return null;
  }

  const containerStyle: CSSProperties = {
    position: "absolute",
    top: overlayPos.top,
    left: overlayPos.left,
    width: overlayPos.width,
    // Match the painted HF rect height so the chrome bar (anchored via
    // `bottom: 100%` for the header, `top: 100%` for the footer) sits flush
    // against the content box instead of collapsing onto the first text row.
    height: overlayPos.height,
    zIndex: 10,
    // No pointer events on the chrome container itself — the painted HF
    // beneath must receive clicks so the pointer pipeline routes them to
    // the persistent HF PM. Individual chrome buttons opt back in.
    pointerEvents: "none",
  };

  return (
    <div
      role="presentation"
      className="hf-inline-editor"
      style={containerStyle}
    >
      {position === "footer" && (
        <div className="hf-separator-bar" style={footerBarStyle}>
          <span style={labelStyle}>{label}</span>
          <OptionsMenu
            position={position}
            showOptions={showOptions}
            setShowOptions={setShowOptions}
            optionsRef={optionsRef}
            onRemove={onRemove}
            onClose={handleClose}
            getActiveView={getActiveView}
          />
        </div>
      )}

      {position === "header" && (
        <div className="hf-separator-bar" style={separatorBarStyle}>
          <span style={labelStyle}>{label}</span>
          <OptionsMenu
            position={position}
            showOptions={showOptions}
            setShowOptions={setShowOptions}
            optionsRef={optionsRef}
            onRemove={onRemove}
            onClose={handleClose}
            getActiveView={getActiveView}
          />
        </div>
      )}
    </div>
  );
});
/* eslint-enable prefer-arrow-callback */

export const InlineHeaderFooterEditor: ForwardRefExoticComponent<
  InlineHeaderFooterEditorProps & RefAttributes<InlineHeaderFooterEditorRef>
> = InlineHeaderFooterEditorComponent;

// ============================================================================
// OPTIONS MENU SUB-COMPONENT
// ============================================================================

function OptionsMenu({
  position,
  showOptions,
  setShowOptions,
  optionsRef,
  onRemove,
  onClose,
  getActiveView,
}: {
  position: "header" | "footer";
  showOptions: boolean;
  setShowOptions: (v: boolean | ((prev: boolean) => boolean)) => void;
  optionsRef: React.RefObject<HTMLDivElement | null>;
  onRemove?: (() => void) | undefined;
  onClose: () => void;
  getActiveView: () => EditorView | null;
}) {
  const t = useTranslations("folio");
  const insertField = (fieldType: "PAGE" | "NUMPAGES") => {
    const view = getActiveView();
    if (!view) {
      return;
    }
    const { $from, from } = view.state.selection;
    const marks = view.state.storedMarks || $from.marks();
    const node = schema.nodes["field"]!.create({
      fieldType,
      instruction: ` ${fieldType} \\* MERGEFORMAT `,
      fieldKind: "simple",
      dirty: true,
    });
    const tr = view.state.tr.insert(from, node.mark(marks));
    view.dispatch(tr);
    view.focus();
  };

  return (
    <div style={{ position: "relative" }} ref={optionsRef}>
      <button
        type="button"
        className="hf-options-button"
        onClick={(e) => {
          e.stopPropagation();
          setShowOptions((prev) => !prev);
        }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        {t("headerFooter.options")} ▾
      </button>
      {showOptions && (
        <div
          className={`hf-options-dropdown${position === "footer" ? " hf-options-dropdown--up" : ""}`}
        >
          <button
            type="button"
            className="hf-options-item"
            onClick={() => {
              setShowOptions(false);
              insertField("PAGE");
            }}
          >
            {t("headerFooter.insertPageNumber")}
          </button>
          <button
            type="button"
            className="hf-options-item"
            onClick={() => {
              setShowOptions(false);
              insertField("NUMPAGES");
            }}
          >
            {t("headerFooter.insertPageCount")}
          </button>
          <div className="hf-options-divider" />
          {onRemove && (
            <button
              type="button"
              className="hf-options-item"
              onClick={() => {
                setShowOptions(false);
                onRemove();
              }}
            >
              {position === "header"
                ? t("headerFooter.removeHeader")
                : t("headerFooter.removeFooter")}
            </button>
          )}
          <button
            type="button"
            className="hf-options-item"
            onClick={() => {
              setShowOptions(false);
              onClose();
            }}
          >
            {position === "header"
              ? t("headerFooter.closeHeaderEditing")
              : t("headerFooter.closeFooterEditing")}
          </button>
        </div>
      )}
    </div>
  );
}
